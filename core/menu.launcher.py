# core/menu_launcher.py
from __future__ import annotations

import os
import sys
import time
import tempfile
import subprocess


_LOCK_NAME = "menu_one.lock"


def _lock_path() -> str:
    return os.path.join(tempfile.gettempdir(), _LOCK_NAME)


def _pid_alive(pid: int) -> bool:
    if pid <= 0:
        return False
    try:
        # Works on Windows too (raises OSError if not alive)
        os.kill(pid, 0)
        return True
    except OSError:
        return False


def _read_lock_pid() -> int:
    try:
        with open(_lock_path(), "r", encoding="utf-8") as f:
            txt = f.read().strip()
        try:
            return int(txt)
        except Exception:
            return -1
    except Exception:
        return -1


def _try_acquire_lock() -> bool:
    """
    Prevents multiple launches if user spams Enter.
    If lock exists but PID is dead -> clears it.
    """
    lp = _lock_path()

    if os.path.exists(lp):
        pid = _read_lock_pid()
        if not _pid_alive(pid):
            try:
                os.remove(lp)
            except Exception:
                pass
        else:
            return False

    # Create lock as "launching"
    try:
        with open(lp, "x", encoding="utf-8") as f:
            f.write("0")
        return True
    except FileExistsError:
        return False
    except Exception:
        return False


def _update_lock_pid(pid: int) -> None:
    try:
        with open(_lock_path(), "w", encoding="utf-8") as f:
            f.write(str(int(pid)))
    except Exception:
        pass


def launch_menu_one() -> bool:
    """
    Launch Menu One as a separate process.
    Returns True if launched, False if already running or failed.
    """
    if not _try_acquire_lock():
        return False

    try:
        # If running as a PyInstaller exe, sys.executable is the exe path
        if getattr(sys, "frozen", False):
            cmd = [sys.executable, "--menu-one"]
        else:
            # Dev mode
            cmd = [sys.executable, "-m", "menu.Menu_One"]

        # Start detached-ish. Use cwd to keep relative logs consistent.
        p = subprocess.Popen(
            cmd,
            cwd=os.getcwd(),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            stdin=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP if os.name == "nt" else 0,
        )

        _update_lock_pid(p.pid)
        return True

    except Exception:
        # If launch fails, remove lock so user isn't stuck
        try:
            os.remove(_lock_path())
        except Exception:
            pass
        return False