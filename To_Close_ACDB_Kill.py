import subprocess

# Kill Access
subprocess.run("taskkill /IM msaccess.exe /F", shell=True)

# Kill orphaned COM host processes
subprocess.run("taskkill /IM dllhost.exe /F", shell=True)
