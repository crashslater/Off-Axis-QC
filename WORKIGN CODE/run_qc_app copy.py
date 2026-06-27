import os
import sys
import time
import threading
import multiprocessing
import subprocess
import webbrowser
import socket
import tempfile
from pathlib import Path
from urllib.request import urlopen
from urllib.error import HTTPError

APP_TITLE = "Off Axis Entertainment GFX QC"

# -----------------------------
# Logging
# -----------------------------
def log_path() -> Path:
    p = Path.home() / "Library" / "Logs" / "OffAxisGFXQC"
    p.mkdir(parents=True, exist_ok=True)
    return p / "launcher.log"

def log(msg: str) -> None:
    try:
        ts = time.strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path(), "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except Exception:
        pass

# -----------------------------
# Dock / Cocoa (no event loop)
# -----------------------------
def setup_dock_presence():
    """
    Make macOS treat this as a normal GUI app (Dock icon persists).
    We do NOT run NSApp.run() because Streamlit needs the main thread.
    """
    if sys.platform != "darwin":
        return
    try:
        from AppKit import NSApplication, NSApplicationActivationPolicyRegular
        app = NSApplication.sharedApplication()
        app.setActivationPolicy_(NSApplicationActivationPolicyRegular)
        app.finishLaunching()
        log("Cocoa app initialized (Dock should persist).")
    except Exception as e:
        log(f"setup_dock_presence failed: {e}")

def stop_dock_bounce():
    """
    Tell macOS we're fully launched/active. This usually stops the Dock icon bouncing.
    Called right before we show the first user-facing dialog.
    """
    if sys.platform != "darwin":
        return
    try:
        from AppKit import NSRunningApplication, NSApplicationActivateIgnoringOtherApps
        NSRunningApplication.currentApplication().activateWithOptions_(
            NSApplicationActivateIgnoringOtherApps
        )
        log("Activated app to stop Dock bounce.")
    except Exception as e:
        log(f"stop_dock_bounce failed: {e}")

# -----------------------------
# PyInstaller resource helper
# -----------------------------
def resource_path(filename: str) -> str:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
    return str(base / filename)

# -----------------------------
# Port + lock
# -----------------------------
def pick_free_port() -> int:
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port

def lock_path() -> Path:
    return Path(tempfile.gettempdir()) / "offaxis_gfxqc.lock"

def _pid_is_running(pid: int) -> bool:
    if pid <= 1:
        return False
    try:
        os.kill(pid, 0)
        return True
    except Exception:
        return False

def acquire_lock_or_exit():
    p = lock_path()

    # Stale lock cleanup
    if p.exists():
        try:
            old_pid = int(p.read_text().strip())
        except Exception:
            old_pid = -1

        if not _pid_is_running(old_pid):
            try:
                p.unlink()
                log(f"Stale lock removed: {p} (pid={old_pid})")
            except Exception as e:
                log(f"Failed to remove stale lock: {p} ({e})")

    try:
        fd = os.open(str(p), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        os.write(fd, str(os.getpid()).encode("utf-8"))
        os.close(fd)
        log(f"Lock acquired: {p}")
        return p
    except FileExistsError:
        log(f"Lock exists (active): {p}")
        subprocess.run(
            ["osascript", "-e",
             f'display dialog "{APP_TITLE} is already running." buttons {{"OK"}} default button "OK"']
        )
        os._exit(0)

def remove_lock(p: Path):
    try:
        os.remove(p)
        log(f"Lock removed: {p}")
    except Exception as e:
        log(f"Lock remove failed: {p} ({e})")

# -----------------------------
# HTTP helpers
# -----------------------------
def http_status(url: str, timeout=0.9) -> int:
    try:
        with urlopen(url, timeout=timeout) as r:
            return r.status
    except HTTPError as e:
        return e.code
    except Exception:
        return 0

def wait_for_health(port: int, timeout_sec: int = 50) -> bool:
    deadline = time.time() + timeout_sec
    health = f"http://127.0.0.1:{port}/_stcore/health"
    while time.time() < deadline:
        hs = http_status(health)
        log(f"Polling health: {hs} ({health})")
        if hs == 200:
            return True
        time.sleep(0.25)
    return False

def resolve_ui_url(port: int) -> str:
    root = f"http://127.0.0.1:{port}/"
    rs = http_status(root)
    log(f"UI probe: {root} -> {rs}")
    return f"http://localhost:{port}/"

# -----------------------------
# Dialog thread
# -----------------------------
def dialog_thread(port: int, lp: Path):
    if not wait_for_health(port, timeout_sec=50):
        health = f"http://127.0.0.1:{port}/_stcore/health"
        hs = http_status(health)

        msg = (
            f"QC server did not become ready.\n\n"
            f"Port: {port}\n"
            f"Health status: {hs}\n\n"
            f"Log: {str(log_path())}\n"
        )
        msg = msg.replace('"', "'")
        log(f"READY TIMEOUT. health={hs}")
        subprocess.run(["osascript", "-e",
                        f'display dialog "{msg}" buttons {{"OK"}} default button "OK"'])
        remove_lock(lp)
        os._exit(1)

    # Stop Dock bounce right when we’re about to show the dialog
    stop_dock_bounce()

    script = f'''
    set choice to button returned of (display dialog "{APP_TITLE} is running." buttons {{"Open QC Tool", "Quit"}} default button "Open QC Tool")
    return choice
    '''
    result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    choice = (result.stdout or "").strip()
    log(f"Dialog choice: {choice}")

    if choice == "Open QC Tool":
        url = resolve_ui_url(port)
        log(f"Opening browser: {url}")
        webbrowser.open(url)
        return

    remove_lock(lp)
    os._exit(0)

# -----------------------------
# Main
# -----------------------------
def main():
    multiprocessing.freeze_support()

    log("=" * 72)
    log(f"Launcher starting. Python={sys.version}")
    log(f"Executable={sys.executable}")
    log(f"MEIPASS={getattr(sys, '_MEIPASS', None)}")

    # Make Dock behave like a normal app
    setup_dock_presence()

    port = pick_free_port()
    log(f"Chosen port: {port}")
    log(f"Log file: {log_path()}")

    # Streamlit env
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"
    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
    os.environ["STREAMLIT_SERVER_FILE_WATCHER_TYPE"] = "none"
    os.environ["STREAMLIT_SERVER_RUN_ON_SAVE"] = "false"
    os.environ["STREAMLIT_SERVER_PORT"] = str(port)
    os.environ["STREAMLIT_SERVER_ADDRESS"] = "127.0.0.1"
    os.environ["STREAMLIT_LOG_LEVEL"] = "warning"
    os.environ["STREAMLIT_SERVER_BASE_URL_PATH"] = ""

    lp = acquire_lock_or_exit()

    # Redirect stdout/stderr into our log (captures Streamlit tracebacks)
    try:
        lf = open(log_path(), "a", encoding="utf-8")
        sys.stdout = lf
        sys.stderr = lf
        print("[launcher] stdout/stderr redirected to launcher.log")
    except Exception as e:
        log(f"Failed to redirect stdout/stderr: {e}")

    # Start dialog watcher thread (Streamlit stays main thread)
    t = threading.Thread(target=dialog_thread, args=(port, lp), daemon=True)
    t.start()

    try:
        import streamlit.web.cli as stcli

        qc_script = resource_path("qc_app.py")
        print("[launcher] qc_script:", qc_script)
        print("[launcher] qc_script exists:", os.path.exists(qc_script))

        if not os.path.exists(qc_script):
            msg = (
                "qc_app.py was not found inside the app bundle.\n\n"
                f"Expected path:\n{qc_script}\n\n"
                f"Log:\n{str(log_path())}\n"
            )
            msg = msg.replace('"', "'")
            subprocess.run(["osascript", "-e",
                            f'display dialog "{msg}" buttons {{"OK"}} default button "OK"'])
            remove_lock(lp)
            os._exit(1)

        sys.argv = [
            "streamlit", "run", qc_script,
            "--server.headless=true",
            "--browser.gatherUsageStats=false",
            "--server.runOnSave=false",
            "--server.fileWatcherType=none",
            "--server.baseUrlPath", "",
        ]
        print("[launcher] Streamlit argv:", sys.argv)

        stcli.main()

    finally:
        remove_lock(lp)

if __name__ == "__main__":
    main()
