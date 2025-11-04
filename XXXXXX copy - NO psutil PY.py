import os
import time
import gc
import pythoncom
import win32com.client as win32

ACCESS_DB = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\fbb.accdb"
GRACE_SECONDS = 5

def run_access_macro_simple(db_path, macro_name, timeout=120):
    pythoncom.CoInitialize()
    acc = None
    success = False
    try:
        print(f"Opening Access and loading DB: {db_path}")
        acc = win32.Dispatch("Access.Application")
        acc.Visible = False
        # Open database
        acc.OpenCurrentDatabase(db_path)
        ran = False
        try:
            print(f"Attempting DoCmd.RunMacro('{macro_name}')")
            acc.DoCmd.RunMacro(macro_name)
            ran = True
            print("Macro executed via DoCmd.RunMacro.")
        except Exception as e:
            print(f"DoCmd.RunMacro failed: {e}")
        if not ran:
            try:
                print(f"Attempting Application.Run('{macro_name}')")
                acc.Run(macro_name)
                ran = True
                print("Macro executed via Application.Run.")
            except Exception as e:
                print(f"Application.Run failed: {e}")
        if not ran:
            raise RuntimeError("Failed to invoke macro via COM.")

        # Optional: wait a bit for the macro to complete (increase if your queries take long)
        wait = 0
        print("Waiting for macro to run (waiting by timeout)...")
        while wait < timeout:
            time.sleep(1)
            wait += 1

        # close DB and quit
        try:
            acc.CloseCurrentDatabase()
        except Exception as e:
            print(f"CloseCurrentDatabase error: {e}")
        try:
            acc.Quit()
        except Exception as e:
            print(f"Quit error: {e}")

        success = True
    except Exception as e:
        print(f"Error during run_access_macro_simple: {e}")
    finally:
        try:
            if acc:
                del acc
        except Exception:
            pass
        pythoncom.CoUninitialize()
    return success

def main():
    ran_ok = run_access_macro_simple(ACCESS_DB, "RunMyQueriesFBB", timeout=10)
    print(f"Macro run status: {'OK' if ran_ok else 'FAILED'}")
    print(f"Waiting {GRACE_SECONDS} seconds for Access to close...")
    time.sleep(GRACE_SECONDS)
    print("Done.")

if __name__ == "__main__":
    main()
