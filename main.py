import sys
import os
import errno
import tempfile
import atexit
import psutil
from CrystalDex import CrystalDex_main
from CrystalDex import update_excel
from CrystalDex import get_runs

LOCK_FILE = os.path.join(tempfile.gettempdir(),"CrystalDex.lock")

def running():
    try:
        fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_RDWR)
        os.write(fd, str(os.getpid()).encode())
        os.close(fd)
        atexit.register(lambda: os.remove(LOCK_FILE))
        return False
    except OSError as e:
        if e.errno == errno.EEXIST:
            try:
                with open(LOCK_FILE, "r") as f:
                    old_pid = int(f.read().strip())
                if psutil.pid_exists(old_pid):
                    return True
                else:
                    os.remove(LOCK_FILE)
                    return running()
            except Exception:
                return True
        raise

if running():
    print("CrystalDex is already running.")
    sys.exit(0)
elif __name__ == '__main__':
    app = CrystalDex_main()
    app.startup()
    app.root.after_idle(app.refocus)
    app.root.mainloop()
    update_excel()
    app.run = get_runs()
    if isinstance(app.run[-1],int):
        app.run_sheet(app.run[-1])