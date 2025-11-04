# force_close_access.py
import os
import sys
import time
import gc
import ctypes
import pythoncom
import pywintypes
import win32com.client

try:
    import psutil
except Exception:
    psutil = None

ACCESS_DB = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\PT\Fil_November_SalesKPITA.accdb"
GRACE_SECONDS = 3
FORCE_KILL = True  # set False to only diagnose and attempt COM shutdowns

def norm_path(p):
    try:
        return os.path.normcase(os.path.abspath(p))
    except Exception:
        return p

def list_msaccess_processes():
    procs = []
    if psutil is None:
        print("ℹ️ psutil not available; cannot list processes in detail")
        return procs
    for proc in psutil.process_iter(["pid", "name", "cmdline", "exe"]):
        try:
            name = (proc.info.get("name") or "").lower()
            if "msaccess" in name or (proc.info.get("exe") or "").lower().endswith("msaccess.exe"):
                procs.append(proc)
        except Exception:
            continue
    return procs

def get_window_title(pid):
    # minimal Win32 enum to find a window title for this pid (may return first top-level)
    EnumWindows = ctypes.windll.user32.EnumWindows
    GetWindowThreadProcessId = ctypes.windll.user32.GetWindowThreadProcessId
    IsWindowVisible = ctypes.windll.user32.IsWindowVisible
    GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
    GetWindowText = ctypes.windll.user32.GetWindowTextW

    titles = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def _enum(hwnd, lparam):
        if not IsWindowVisible(hwnd):
            return True
        pid = ctypes.c_ulong()
        GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if pid.value == lparam:
            length = GetWindowTextLength(hwnd)
            if length > 0:
                buf = ctypes.create_unicode_buffer(length + 1)
                GetWindowText(hwnd, buf, length + 1)
                titles.append(buf.value)
        return True

    EnumWindows(_enum, pid)
    return titles

def try_getobject_quit(db_path):
    try:
        print("1) Trying win32com.client.GetObject(db_path)...")
        app = win32com.client.GetObject(db_path)
        if app is None:
            print("  → GetObject returned None")
            return False
        try:
            app.Quit()
            print("  → Called Quit() on attached Application")
        except Exception as ex:
            print("  → Quit() raised:", ex)
            try:
                if hasattr(app, "CloseCurrentDatabase"):
                    app.CloseCurrentDatabase()
                    print("  → Called CloseCurrentDatabase()")
                app.Quit()
                print("  → Called Quit() after CloseCurrentDatabase()")
            except Exception as ex2:
                print("  → Secondary quit attempt failed:", ex2)
        try:
            del app
        except Exception:
            pass
        gc.collect()
        return True
    except pywintypes.com_error as exc:
        print("  → GetObject COM error:", exc)
        return False
    except Exception as exc:
        print("  → GetObject failed:", exc)
        return False

def try_getactive_release(db_path):
    try:
        print("2) Trying GetActiveObject('Access.Application')...")
        app = win32com.client.GetActiveObject("Access.Application")
    except pywintypes.com_error:
        print("  → No active Access.Application found in ROT")
        return False
    except Exception as ex:
        print("  → GetActiveObject failed:", ex)
        return False

    try:
        cp = getattr(app, "CurrentProject", None)
        full = getattr(cp, "FullName", None) if cp is not None else None
        print("  → Active object's CurrentProject.FullName:", full)
        # attempt to quit regardless of match to release orphaned instance
        try:
            app.Quit()
            print("  → Called Quit() on ActiveObject")
        except Exception as ex:
            print("  → Quit() on ActiveObject raised:", ex)
        try:
            del app
        except Exception:
            pass
        gc.collect()
        return True
    except Exception as ex:
        print("  → Error inspecting ActiveObject:", ex)
        try:
            del app
        except Exception:
            pass
        gc.collect()
        return False

def terminate_process(proc, reason):
    try:
        print(f"  → Terminating PID {proc.pid} ({reason})")
        proc.terminate()
    except Exception as ex:
        print(f"  → terminate() failed for PID {proc.pid}:", ex)
    # wait short grace
    try:
        proc.wait(timeout=GRACE_SECONDS)
        print(f"  → PID {proc.pid} exited after terminate()")
        return True
    except Exception:
        try:
            print(f"  → Killing PID {proc.pid}")
            proc.kill()
            proc.wait(timeout=5)
            print(f"  → PID {proc.pid} killed")
            return True
        except Exception as ex:
            print(f"  → kill() failed for PID {proc.pid}:", ex)
            return False

def force_kill_matching(db_path):
    if psutil is None:
        print("ℹ️ psutil not installed; cannot force-kill processes. Install with: pip install psutil")
        return False
    normalized_target = norm_path(db_path)
    any_found = False
    procs = list_msaccess_processes()
    print(f"Found {len(procs)} msaccess.exe process(es).")
    for p in procs:
        try:
            cmd = " ".join(p.info.get("cmdline") or [])
            exe = p.info.get("exe") or ""
            window_titles = get_window_title(p.pid)
            title_summary = ";".join(window_titles) if window_titles else ""
            matches_db = normalized_target in os.path.normcase(cmd) or normalized_target in os.path.normcase(exe)
            print(f" PID {p.pid} | exe: {exe} | cmdline contains DB: {matches_db} | window titles: {title_summary}")
            if matches_db or not window_titles:
                any_found = True
                if FORCE_KILL:
                    ok = terminate_process(p, "matching DB or no visible window")
                    if not ok:
                        print(f"  → Failed to terminate PID {p.pid}")
                else:
                    print(f"  → Would terminate PID {p.pid} (FORCE_KILL=False)")
        except Exception as ex:
            print("  → Error examining process:", ex)
            continue
    return any_found

def main():
    print("=== Diagnostic: list Access processes and try COM shutdowns ===")
    print("Target DB:", ACCESS_DB)
    print()

    # 0) list processes before
    procs = list_msaccess_processes()
    for p in procs:
        try:
            cmd = " ".join(p.info.get("cmdline") or [])
            exe = p.info.get("exe") or ""
            titles = get_window_title(p.pid)
            print(f"Before: PID {p.pid} | exe:{exe} | cmd:{cmd} | titles: {titles}")
        except Exception:
            continue

    # Initialize COM for our calls
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

    # 1) try GetObject attach+Quit
    try_getobject_quit(ACCESS_DB)
    gc.collect()
    time.sleep(0.5)

    # 2) try GetActiveObject and Quit
    try_getactive_release(ACCESS_DB)
    gc.collect()
    time.sleep(0.5)

    # Allow Access some time to exit cleanly
    print(f"Waiting up to {GRACE_SECONDS} seconds for Access to close cleanly...")
    time.sleep(GRACE_SECONDS)

    # 3) inspect processes and force kill if necessary
    print()
    print("Inspecting processes after COM quit attempts...")
    any_matched = force_kill_matching(ACCESS_DB)

    # 4) final list
    print()
    print("Final msaccess.exe processes:")
    final = list_msaccess_processes()
    for p in final:
        try:
            cmd = " ".join(p.info.get("cmdline") or [])
            exe = p.info.get("exe") or ""
            titles = get_window_title(p.pid)
            print(f"After: PID {p.pid} | exe:{exe} | cmd:{cmd} | titles: {titles}")
        except Exception:
            continue

    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass

    if any_matched:
        print("✅ Attempts made to terminate matching msaccess processes.")
    else:
        print("ℹ️ No matching ms access processes found for the target DB. If Access still appears as a zombie, it may be a hidden COM server or a different user session.")

if __name__ == "__main__":
    main()
