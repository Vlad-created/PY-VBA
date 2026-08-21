
import argparse
import functools
import logging
import os
import subprocess
import sys
import time
from typing import Callable, Iterable, List, Optional

import pythoncom
import pywintypes
import win32com.client as win32
from win32com.client import DispatchEx

# -------------------------
# Configuration / constants
# -------------------------
REFRESH_TIMEOUT_SECONDS = int(os.environ.get("REFRESH_TIMEOUT_SECONDS", 3600))
POLL_INTERVAL = float(os.environ.get("POLL_INTERVAL", 2.0))
HEARTBEAT_SECONDS = float(os.environ.get("HEARTBEAT_SECONDS", 10.0))
EXCEL_CLOSE_WAIT_SECONDS = float(os.environ.get("EXCEL_CLOSE_WAIT_SECONDS", 5.0))
RETRY_ATTEMPTS = int(os.environ.get("RETRY_ATTEMPTS", 3))
RETRY_BACKOFF = float(os.environ.get("RETRY_BACKOFF", 2.0))  # multiplier

# Excel constants (kept local for clarity)
xlDone = 0
xlLinkTypeExcelLinks = 1

# COM/Excel busy error
RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846
RETRYABLE_COM_HRESULTS = {
    RPC_E_CALL_REJECTED,
    RPC_E_SERVERCALL_RETRYLATER,
}

EXCEL_READY_TIMEOUT_SECONDS = float(
    os.environ.get("EXCEL_READY_TIMEOUT_SECONDS", 120.0)
)
COM_RETRY_ATTEMPTS = int(
    os.environ.get("COM_RETRY_ATTEMPTS", 60)
)
COM_RETRY_DELAY_SECONDS = float(
    os.environ.get("COM_RETRY_DELAY_SECONDS", 1.0)
)


def _com_hresult(exc):
    try:
        return exc.args[0]
    except Exception:
        return None


def wait_for_excel_ready(
    excel,
    timeout=EXCEL_READY_TIMEOUT_SECONDS,
    poll_interval=1.0,
):
    start = time.time()

    while time.time() - start < timeout:
        try:
            ready = bool(excel.Ready)
        except pywintypes.com_error as e:
            if _com_hresult(e) in RETRYABLE_COM_HRESULTS:
                ready = False
            else:
                raise
        except Exception:
            ready = False

        if ready:
            try:
                if excel.CalculationState == xlDone:
                    return True
            except pywintypes.com_error as e:
                if _com_hresult(e) not in RETRYABLE_COM_HRESULTS:
                    raise
            except Exception:
                pass

        try:
            pythoncom.PumpWaitingMessages()
        except Exception:
            pass

        time.sleep(poll_interval)

    log(
        f"WARNING: Excel was not Ready within "
        f"{timeout:.0f}s. Continuing.",
        "warning",
    )
    return False


def safe_com_call(
    func,
    description,
    attempts=COM_RETRY_ATTEMPTS,
    delay=COM_RETRY_DELAY_SECONDS,
):
    last_error = None

    for attempt in range(1, attempts + 1):
        try:
            return func()

        except pywintypes.com_error as e:
            last_error = e
            hresult = _com_hresult(e)

            if hresult not in RETRYABLE_COM_HRESULTS:
                raise

            if attempt == 1:
                log(
                    f"  Excel is busy while {description}. "
                    f"Retrying..."
                    ,
                    "warning",
                )
            elif attempt % 10 == 0:
                log(
                    f"  Still waiting for Excel while {description} "
                    f"({attempt}s)...",
                    "warning",
                )

            try:
                pythoncom.PumpWaitingMessages()
            except Exception:
                pass

            time.sleep(delay)

        except Exception:
            raise

    raise RuntimeError(
        f"Excel remained busy while {description} "
        f"after {attempts} attempts."
    ) from last_error


def refresh_one_connection(
    conn,
    wb,
    excel,
    timeout=REFRESH_TIMEOUT_SECONDS,
):
    name = str(getattr(conn, "Name", "unknown"))

    wait_for_excel_ready(excel)

    log(f"  Refreshing: {name}")

    safe_com_call(
        conn.Refresh,
        f"refreshing connection '{name}'",
    )

    wait_for_connections(
        wb,
        excel,
        connections=[conn],
        timeout=timeout,
    )

    wait_for_excel_ready(excel)

    try:
        safe_com_call(
            excel.CalculateUntilAsyncQueriesDone,
            f"waiting for asynchronous queries after '{name}'",
            attempts=30,
            delay=1.0,
        )
    except Exception as e:
        log(
            f"  WARNING: CalculateUntilAsyncQueriesDone "
            f"was not available/completed after '{name}': {e}",
            "warning",
        )

    wait_for_excel_ready(excel)

LOG_FILE = os.environ.get("REFRESH_LOG_FILE", "refresh_workbooks.log")

# -------------------------
# Logging
# -------------------------
logger = logging.getLogger("refresh")
logger.setLevel(logging.INFO)
formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")

# Console handler
ch = logging.StreamHandler(sys.stdout)
ch.setFormatter(formatter)
logger.addHandler(ch)

# File handler
fh = logging.FileHandler(LOG_FILE)
fh.setFormatter(formatter)
logger.addHandler(fh)


def log(msg: str, level: str = "info"):
    if level == "debug":
        logger.debug(msg)
    elif level == "warning":
        logger.warning(msg)
    elif level == "error":
        logger.error(msg)
    else:
        logger.info(msg)


# -------------------------
# Utilities
# -------------------------
def retry(
    attempts: int = RETRY_ATTEMPTS,
    backoff: float = RETRY_BACKOFF,
    exceptions: tuple = (Exception,),
):
    """Retry decorator with exponential backoff."""
    def deco(func: Callable):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            last_exc = None
            delay = 1.0
            for attempt in range(1, attempts + 1):
                try:
                    return func(*args, **kwargs)
                except exceptions as e:
                    last_exc = e
                    log(f"Attempt {attempt}/{attempts} failed for {func.__name__}: {e}", "warning")
                    if attempt < attempts:
                        time.sleep(delay)
                        delay *= backoff
            log(f"All {attempts} attempts failed for {func.__name__}: {last_exc}", "error")
            raise last_exc
        return wrapper
    return deco


def safe_quit_excel(excel_app):
    if excel_app is None:
        return

    quit_ok = False

    for attempt in range(1, COM_RETRY_ATTEMPTS + 1):
        try:
            excel_app.Quit()
            quit_ok = True
            break

        except pywintypes.com_error as e:
            if _com_hresult(e) not in RETRYABLE_COM_HRESULTS:
                log(f"Exception while quitting Excel: {e}", "warning")
                break

            if attempt == 1:
                log(
                    "Excel is still busy while closing. "
                    "Retrying...",
                    "warning",
                )
            elif attempt % 10 == 0:
                log(
                    f"  Still waiting to close Excel "
                    f"({attempt * COM_RETRY_DELAY_SECONDS:.0f}s)...",
                    "warning",
                )

            try:
                pythoncom.PumpWaitingMessages()
            except Exception:
                pass

            time.sleep(COM_RETRY_DELAY_SECONDS)

        except Exception as e:
            log(f"Exception while quitting Excel: {e}", "warning")
            break

    if quit_ok:
        log("Excel closed.")
    else:
        log(
            "WARNING: Excel.Quit() could not be completed.",
            "warning",
        )

    log(
        f"Waiting {EXCEL_CLOSE_WAIT_SECONDS} seconds "
        f"before next Excel file..."
    )
    time.sleep(EXCEL_CLOSE_WAIT_SECONDS)


# -------------------------
# Access macros
# -------------------------
@retry()
def run_access_macros(db_path: str, macro_names: Iterable[str]):
    log(f"\nOpening database:\n{db_path}")
    pythoncom.CoInitialize()
    access = None
    try:
        access = DispatchEx("Access.Application")
        access.Visible = False
        access.OpenCurrentDatabase(db_path)
        for macro_name in macro_names:
            log(f"Running macro: {macro_name}")
            access.DoCmd.RunMacro(macro_name)
            log(f"  -> '{macro_name}' completed.")
        access.CloseCurrentDatabase()
    finally:
        if access is not None:
            try:
                access.Quit()
            except Exception as e:
                log(f"Error quitting Access: {e}", "warning")
            del access
        pythoncom.CoUninitialize()
        log("Database closed.")


# -------------------------
# Refresh monitoring helpers
# -------------------------
def any_refreshing(workbook, app, connections=None):
    """Return True if any connection/pivot/calculation is still running."""
    conns = connections if connections is not None else getattr(workbook, "Connections", None)

    if conns is not None:
        try:
            for conn in conns:
                try:
                    if getattr(conn, "Refreshing", False):
                        return True
                except Exception:
                    # ignore connection-specific errors
                    pass
        except Exception:
            pass

    try:
        if getattr(app, "CalculationState", xlDone) != xlDone:
            return True
    except Exception:
        pass

    # Some workbooks expose a Refreshing property
    try:
        if getattr(workbook, "Refreshing", False):
            return True
    except Exception:
        pass

    return False


def wait_for_connections(wb, excel, connections=None, timeout=REFRESH_TIMEOUT_SECONDS):
    start = time.time()
    last_heartbeat = start

    while any_refreshing(wb, excel, connections=connections):
        elapsed = time.time() - start
        if elapsed > timeout:
            log("WARNING: refresh timeout reached. Proceeding anyway.", "warning")
            break
        if time.time() - last_heartbeat > HEARTBEAT_SECONDS:
            log(f"  ...still refreshing ({int(elapsed)}s)")
            last_heartbeat = time.time()
        time.sleep(POLL_INTERVAL)
    else:
        log(f"Refresh finished in {int(time.time() - start)}s.")


# -------------------------
# Pivot refresh
# -------------------------
def refresh_all_pivots(wb):
    pivot_count = 0
    for sheet in wb.Sheets:
        try:
            pts = sheet.PivotTables()
        except Exception:
            continue
        for i in range(1, pts.Count + 1):
            pt = pts.Item(i)
            try:
                log(f"  Refreshing pivot '{pt.Name}' on sheet '{sheet.Name}'...")
                safe_com_call(
                    pt.RefreshTable,
                    f"refreshing pivot '{pt.Name}'",
                )
                pivot_count += 1
            except Exception as e:
                log(
                    f"  ERROR: failed to refresh pivot "
                    f"'{getattr(pt, 'Name', 'unknown')}': {e}",
                    "error",
                )
                raise
    log(f"Refreshed {pivot_count} pivot table(s).")


# -------------------------
# External links
# -------------------------
def update_workbook_links(wb):
    try:
        links = wb.LinkSources(xlLinkTypeExcelLinks)
    except Exception:
        links = None

    if not links:
        log("No external workbook links found.")
        return

    for link in links:
        log(f"  Updating workbook link: {link}")
        safe_com_call(
            lambda: wb.UpdateLink(
                Name=link,
                Type=xlLinkTypeExcelLinks,
            ),
            f"updating workbook link '{link}'",
        )
    log("Workbook link(s) updated.")


# -------------------------
# Standard Excel refresh
# -------------------------
@retry()
def refresh_excel_queries(xlsx_path: str, save: bool = True, timeout: int = REFRESH_TIMEOUT_SECONDS, refresh_pivots: bool = False):
    log(f"\nOpening workbook:\n{xlsx_path}")
    pythoncom.CoInitialize()
    excel = None
    wb = None
    saved = False
    try:
        excel = DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False

        wb = excel.Workbooks.Open(xlsx_path, UpdateLinks=0)

        # disable background queries where possible
        for conn in wb.Connections:
            try:
                conn.OLEDBConnection.BackgroundQuery = False
            except Exception:
                pass
            try:
                conn.ODBCConnection.BackgroundQuery = False
            except Exception:
                pass

        log("Refreshing all connections/queries...")
        safe_com_call(
            wb.RefreshAll,
            "refreshing all connections/queries",
        )
        wait_for_connections(wb, excel, timeout=timeout)
        wait_for_excel_ready(excel)

        if refresh_pivots:
            log("Refreshing pivot tables...")
            wait_for_excel_ready(excel)
            refresh_all_pivots(wb)
            wait_for_connections(wb, excel, timeout=timeout)
            wait_for_excel_ready(excel)

        if save:
            safe_com_call(
                wb.Save,
                "saving workbook",
            )
            saved = True
            log("Workbook saved.")

        safe_com_call(
            lambda: wb.Close(SaveChanges=False),
            "closing workbook",
            attempts=30,
            delay=1.0,
        )
        wb = None
        log("Workbook closed.")
    finally:
        # ensure workbook reference removed
        try:
            if wb is not None:
                del wb
        except Exception:
            pass
        if excel is not None:
            safe_quit_excel(excel)
            try:
                del excel
            except Exception:
                pass
        pythoncom.CoUninitialize()


# -------------------------
# Base connection -> queries
# -------------------------
@retry()
def refresh_connection_then_queries(
    xlsx_path: str,
    save: bool = True,
    timeout: int = REFRESH_TIMEOUT_SECONDS,
    refresh_pivots: bool = True,
):
    log(f"\nOpening workbook:\n{xlsx_path}")

    pythoncom.CoInitialize()
    excel = None
    wb = None

    try:
        excel = DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False

        wb = excel.Workbooks.Open(
            xlsx_path,
            UpdateLinks=0,
        )

        for conn in wb.Connections:
            try:
                conn.OLEDBConnection.BackgroundQuery = False
            except Exception:
                pass

            try:
                conn.ODBCConnection.BackgroundQuery = False
            except Exception:
                pass

        base_conns = []
        query_conns = []

        for conn in wb.Connections:
            name = str(conn.Name)

            if (
                name.startswith("Query -")
                or name.startswith("Query ")
            ):
                query_conns.append(conn)
            else:
                base_conns.append(conn)

        if not base_conns or not query_conns:
            log(
                "Could not distinguish base connection "
                "vs queries."
            )
            log("Refreshing everything together.")

            safe_com_call(
                wb.RefreshAll,
                "refreshing all workbook connections",
            )

            wait_for_connections(
                wb,
                excel,
                timeout=timeout,
            )

            wait_for_excel_ready(excel)

        else:
            log(
                f"Refreshing base connection(s): "
                f"{[c.Name for c in base_conns]}"
            )

            for conn in base_conns:
                refresh_one_connection(
                    conn,
                    wb,
                    excel,
                    timeout=timeout,
                )

            log(
                f"Refreshing {len(query_conns)} "
                f"dependent quer(y/ies): "
                f"{[c.Name for c in query_conns]}"
            )

            for conn in query_conns:
                refresh_one_connection(
                    conn,
                    wb,
                    excel,
                    timeout=timeout,
                )

            log("All dependent queries completed.")

        if refresh_pivots:
            wait_for_excel_ready(excel)

            log("Refreshing pivot tables...")
            refresh_all_pivots(wb)

            wait_for_connections(
                wb,
                excel,
                timeout=timeout,
            )

            wait_for_excel_ready(excel)

        if save:
            safe_com_call(
                wb.Save,
                "saving workbook",
            )
            log("Workbook saved.")

        safe_com_call(
            lambda: wb.Close(SaveChanges=False),
            "closing workbook",
            attempts=30,
            delay=1.0,
        )
        wb = None
        log("Workbook closed.")

    finally:
        if wb is not None:
            try:
                safe_com_call(
                    lambda: wb.Close(SaveChanges=False),
                    "closing workbook after error",
                    attempts=30,
                    delay=1.0,
                )
            except Exception as e:
                log(
                    f"WARNING: workbook close after error failed: {e}",
                    "warning",
                )

        if excel is not None:
            safe_quit_excel(excel)

        pythoncom.CoUninitialize()


# -------------------------
# Links -> queries
# -------------------------
@retry()
def refresh_links_then_queries(xlsx_path: str, save: bool = True, timeout: int = REFRESH_TIMEOUT_SECONDS, refresh_pivots: bool = False):
    log(f"\nOpening workbook:\n{xlsx_path}")
    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False

        wb = excel.Workbooks.Open(xlsx_path, UpdateLinks=0)

        log("Updating workbook link(s)...")
        update_workbook_links(wb)

        for conn in wb.Connections:
            try:
                conn.OLEDBConnection.BackgroundQuery = False
            except Exception:
                pass
            try:
                conn.ODBCConnection.BackgroundQuery = False
            except Exception:
                pass

        log("Refreshing all queries...")
        safe_com_call(
            wb.RefreshAll,
            "refreshing all queries",
        )
        wait_for_connections(wb, excel, timeout=timeout)
        wait_for_excel_ready(excel)

        if refresh_pivots:
            log("Refreshing pivot tables...")
            wait_for_excel_ready(excel)
            refresh_all_pivots(wb)
            wait_for_connections(wb, excel, timeout=timeout)
            wait_for_excel_ready(excel)

        if save:
            safe_com_call(
                wb.Save,
                "saving workbook",
            )
            log("Workbook saved.")

        safe_com_call(
            lambda: wb.Close(SaveChanges=False),
            "closing workbook",
            attempts=30,
            delay=1.0,
        )
        wb = None
        log("Workbook closed.")
    finally:
        if wb is not None:
            try:
                safe_com_call(
                    lambda: wb.Close(SaveChanges=False),
                    "closing workbook after error",
                    attempts=30,
                    delay=1.0,
                )
            except Exception as e:
                log(
                    f"WARNING: workbook close after error failed: {e}",
                    "warning",
                )
        try:
            if wb is not None:
                del wb
        except Exception:
            pass
        if excel is not None:
            safe_quit_excel(excel)
            try:
                del excel
            except Exception:
                pass
        pythoncom.CoUninitialize()


# -------------------------
# Power BI helpers
# -------------------------
def find_powerbi_connections(wb):
    powerbi_connections = []
    for conn in wb.Connections:
        try:
            if not hasattr(conn, "OLEDBConnection"):
                continue
            cs = str(conn.OLEDBConnection.Connection or "").lower()
            if ("msolap" in cs) or ("pbiazure" in cs) or ("powerbi" in cs) or ("claimstoken" in cs):
                powerbi_connections.append(conn)
        except Exception:
            pass
    return powerbi_connections


def wait_for_powerbi_authentication(interactive_hint: Optional[str] = None):
    print()
    print("=" * 70)
    print("POWER BI / MICROSOFT AUTHENTICATION REQUIRED")
    print("=" * 70)
    print("Complete the Microsoft authentication window in Excel.")
    if interactive_hint:
        print(interactive_hint)
    print("Then return here and press ENTER.")
    print("=" * 70)

    try:
        input("Press ENTER after authentication is complete...")
    except Exception:
        log("Interactive input unavailable. Waiting 60 seconds...", "warning")
        time.sleep(60)
    log("Microsoft / Power BI authentication confirmed.")


def find_dependent_query_connections(wb, excluded_connections=None):
    if excluded_connections is None:
        excluded_connections = []
    excluded_ids = {id(conn) for conn in excluded_connections}
    query_names = []
    try:
        for query in wb.Queries:
            try:
                query_names.append(str(query.Name))
            except Exception:
                pass
    except Exception:
        pass

    query_connections = []
    for conn in wb.Connections:
        try:
            if id(conn) in excluded_ids:
                continue
            conn_name = str(conn.Name)
            is_query_connection = conn_name.startswith("Query -") or conn_name.startswith("Query ")
            matches_query_name = any(
                conn_name == q or conn_name == f"Query - {q}" or conn_name == f"Query {q}"
                for q in query_names
            )
            if is_query_connection or matches_query_name:
                query_connections.append(conn)
        except Exception:
            pass
    return query_connections


# -------------------------
# Power BI flows
# -------------------------
@retry()
def refresh_pbi_calculator(xlsx_path: str, save: bool = True, timeout: int = REFRESH_TIMEOUT_SECONDS):
    log(f"\nOpening Power BI workbook:\n{xlsx_path}")
    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False

        wb = excel.Workbooks.Open(xlsx_path, UpdateLinks=0)
        log("Workbook opened.")

        powerbi_connections = find_powerbi_connections(wb)
        if not powerbi_connections:
            raise RuntimeError("Power BI / MSOLAP connection not found.")
        log(f"Power BI connection(s): {[c.Name for c in powerbi_connections]}")

        for conn in wb.Connections:
            try:
                conn.OLEDBConnection.BackgroundQuery = False
            except Exception:
                pass
            try:
                conn.ODBCConnection.BackgroundQuery = False
            except Exception:
                pass

        # interactive auth required
        if os.environ.get("PBI_NONINTERACTIVE_FAIL", "0") == "1":
            raise RuntimeError("Power BI authentication required but non-interactive mode is set.")
        wait_for_powerbi_authentication()

        log("Refreshing Power BI connection...")
        for conn in powerbi_connections:
            refresh_one_connection(
                conn,
                wb,
                excel,
                timeout=timeout,
            )

        wait_for_excel_ready(excel)
        log("Power BI refresh completed.")

        query_connections = find_dependent_query_connections(wb, excluded_connections=powerbi_connections)
        if not query_connections:
            raise RuntimeError("Dependent query connection not found.")
        log(f"Dependent query: {[c.Name for c in query_connections]}")

        for conn in query_connections:
            refresh_one_connection(
                conn,
                wb,
                excel,
                timeout=timeout,
            )

        wait_for_excel_ready(excel)
        log("Dependent query refresh completed.")

        wait_for_excel_ready(excel)
        log("Refreshing ALL PivotTables...")
        refresh_all_pivots(wb)
        wait_for_connections(wb, excel, timeout=timeout)
        wait_for_excel_ready(excel)

        if save:
            try:
                safe_com_call(
                    wb.Save,
                    "saving workbook",
                )
                log("Workbook saved.")
            except Exception as e:
                log(f"ERROR saving workbook: {e}", "error")

        safe_com_call(
            lambda: wb.Close(SaveChanges=False),
            "closing workbook",
            attempts=30,
            delay=1.0,
        )
        wb = None
        log("Workbook closed.")
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        if excel is not None:
            safe_quit_excel(excel)
        pythoncom.CoUninitialize()


@retry()
def refresh_pbi_execution(xlsx_path: str, save: bool = True, timeout: int = REFRESH_TIMEOUT_SECONDS):
    log(f"\nOpening Power BI workbook:\n{xlsx_path}")
    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False

        wb = excel.Workbooks.Open(xlsx_path, UpdateLinks=0)
        log("Workbook opened.")

        powerbi_connections = find_powerbi_connections(wb)
        if not powerbi_connections:
            raise RuntimeError("Power BI / MSOLAP connection not found.")
        log(f"Power BI connection(s): {[c.Name for c in powerbi_connections]}")

        for conn in wb.Connections:
            try:
                conn.OLEDBConnection.BackgroundQuery = False
            except Exception:
                pass
            try:
                conn.ODBCConnection.BackgroundQuery = False
            except Exception:
                pass

        if os.environ.get("PBI_NONINTERACTIVE_FAIL", "0") == "1":
            raise RuntimeError("Power BI authentication required but non-interactive mode is set.")
        wait_for_powerbi_authentication()

        log("Refreshing Power BI connection...")
        for conn in powerbi_connections:
            refresh_one_connection(
                conn,
                wb,
                excel,
                timeout=timeout,
            )

        wait_for_excel_ready(excel)
        log("Power BI refresh completed.")

        log("Updating workbook links...")
        update_workbook_links(wb)

        wait_for_excel_ready(excel)
        log("Refreshing ALL PivotTables...")
        refresh_all_pivots(wb)
        wait_for_connections(wb, excel, timeout=timeout)
        wait_for_excel_ready(excel)

        if save:
            try:
                safe_com_call(
                    wb.Save,
                    "saving workbook",
                )
                log("Workbook saved.")
            except Exception as e:
                log(f"ERROR saving workbook: {e}", "error")

        safe_com_call(
            lambda: wb.Close(SaveChanges=False),
            "closing workbook",
            attempts=30,
            delay=1.0,
        )
        wb = None
        log("Workbook closed.")
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        if excel is not None:
            safe_quit_excel(excel)
        pythoncom.CoUninitialize()


# -------------------------
# Top-level runner
# -------------------------
def maybe_taskkill(force_kill: bool):
    if not force_kill:
        log("Skipping taskkill of msaccess.exe and excel.exe (use --force-kill to enable).")
        return
    log("Running taskkill for msaccess.exe and excel.exe (force).")
    try:
        subprocess.call("taskkill /f /im msaccess.exe", shell=True)
        subprocess.call("taskkill /f /im excel.exe", shell=True)
    except Exception as e:
        log(f"taskkill failed: {e}", "warning")


def main(dry_run: bool = False, force_kill: bool = False):
    maybe_taskkill(force_kill)

    # Example sequence of tasks (preserve your original order)
    # Edit TASKS below to change files/macros or add new steps.
    failed_steps = []

    for step in TASKS:
        name = step.get("name")
        func = step.get("func")
        args = step.get("args", [])
        kwargs = step.get("kwargs", {})

        log(f"\n=== Step: {name} ===")
        if dry_run:
            log(
                f"DRY RUN: would call "
                f"{func.__name__}({', '.join(map(str, args))}, {kwargs})"
            )
            continue

        try:
            func(*args, **kwargs)
        except Exception as e:
            failed_steps.append((name, e))
            log(f"ERROR in step '{name}': {e}", "error")

    if failed_steps:
        log(
            f"\nRun completed with {len(failed_steps)} failed step(s).",
            "error",
        )
        for name, error in failed_steps:
            log(f"FAILED: {name} -> {error}", "error")
        raise RuntimeError(
            f"{len(failed_steps)} task(s) failed. "
            "See log for details."
        )

    log("\nAll processes completed successfully.")


# -------------------------
# TASKS: edit these to match your environment
# -------------------------
# Keep the same UNC paths and macros you had originally; update as needed.
TASKS = [
    {
        "name": "ACCESS - GOAT_FTW.accdb",
        "func": run_access_macros,
        "args": [r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\GOAT_FTW.accdb", ["Macro1 - current month"]],
    },
    {
        "name": "ACCESS - Rank-ACDB.accdb",
        "func": run_access_macros,
        "args": [r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\Rank-ACDB.accdb", ["Macro1"]],
    },
    {
        "name": "EXCEL - Users_ALL_List_TW.xlsx",
        "func": refresh_excel_queries,
        "args": [r"\\virgo\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\1.Users_ALL_List_TW.xlsx"],
        "kwargs": {"refresh_pivots": False},
    },
    {
        "name": "EXCEL - Pivot_Gas_and_Sales.xlsm",
        "func": refresh_excel_queries,
        "args": [r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\Pivot_Gas_and_Sales.xlsm"],
        "kwargs": {"refresh_pivots": True},
    },
    {
        "name": "EXCEL - GOAT_FTW_backup.xlsx (base->queries)",
        "func": refresh_connection_then_queries,
        "args": [r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\GOAT_FTW_backup.xlsx"],
        "kwargs": {"refresh_pivots": True},
    },
    {
        "name": "EXCEL - Pivot_SC.xlsx",
        "func": refresh_excel_queries,
        "args": [r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\Pivot_SC.xlsx"],
        "kwargs": {"refresh_pivots": True},
    },
    {
        "name": "EXCEL - RollOut.xlsx (links->queries)",
        "func": refresh_links_then_queries,
        "args": [r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\RollOut.xlsx"],
    },
    {
        "name": "EXCEL - STAFF.xlsx",
        "func": refresh_excel_queries,
        "args": [r"\\virgo\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\STAFF.xlsx"],
        "kwargs": {"refresh_pivots": False},
    },
    {
        "name": "PBI CALCULATOR - Pivot_live_CALCULATOR.xlsx",
        "func": refresh_pbi_calculator,
        "args": [r"\\virgo\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\Pivot_live_CALCULATOR.xlsx"],
    },
    {
        "name": "PBI EXECUTION - Pivot_live_EXECUTION.xlsx",
        "func": refresh_pbi_execution,
        "args": [r"\\virgo\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\Pivot_live_EXECUTION.xlsx"],
    },
]


# -------------------------
# CLI
# -------------------------
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Refresh Excel/Access/Power BI workbooks.")
    parser.add_argument("--dry-run", action="store_true", help="Show planned tasks without executing.")
    parser.add_argument("--force-kill", action="store_true", help="Run taskkill for msaccess.exe and excel.exe at start (dangerous).")
    args = parser.parse_args()

    try:
        main(dry_run=args.dry_run, force_kill=args.force_kill)
    except Exception as e:
        log(f"Fatal error: {e}", "error")
        sys.exit(1)
