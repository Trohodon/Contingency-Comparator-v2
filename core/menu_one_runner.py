# core/menu_one_runner.py
#
# Launch + single-instance guard for the Menu One easter egg.
# Works in dev (python) and in PyInstaller onefile EXE.

from __future__ import annotations

import atexit
import os
import sys
import tempfile
import subprocess

# IMPORTANT: This import ensures PyInstaller includes the menu package.
# (Even if you only run it via a flag in the EXE)
try:
    import menu.Menu_One as _menu_one  # noqa: F401
except Exception:
    _menu_one = None  # will still work in frozen EXE if bundled correctly


_LOCK_FH = None


def _lockfile_path() -> str:
    # per-user temp lock
    return os.path.join(tempfile.gettempdir(), "menu_one.lock")


def _try_acquire_lock() -> bool:
    """
    Cross-process lock.
    - Windows: msvcrt.locking
    - Unix: fcntl.flock
    """
    global _LOCK_FH
    if _LOCK_FH is not None:
        return True  # already locked in this process

    path = _lockfile_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)

    try:
        fh = open(path, "a+")
    except Exception:
        return False

    try:
        if os.name == "nt":
            import msvcrt
            try:
                # lock 1 byte
                msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
            except OSError:
                fh.close()
                return False
        else:
            import fcntl
            try:
                fcntl.flock(fh.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            except OSError:
                fh.close()
                return False

        _LOCK_FH = fh
        atexit.register(_release_lock)
        return True

    except Exception:
        try:
            fh.close()
        except Exception:
            pass
        return False


def _release_lock():
    global _LOCK_FH
    if _LOCK_FH is None:
        return
    try:
        if os.name == "nt":
            import msvcrt
            try:
                _LOCK_FH.seek(0)
                msvcrt.locking(_LOCK_FH.fileno(), msvcrt.LK_UNLCK, 1)
            except Exception:
                pass
        else:
            import fcntl
            try:
                fcntl.flock(_LOCK_FH.fileno(), fcntl.LOCK_UN)
            except Exception:
                pass
    finally:
        try:
            _LOCK_FH.close()
        except Exception:
            pass
        _LOCK_FH = None


def launch_menu_one_detached(root_dir: str | None = None) -> bool:
    """
    Launch Menu One as a separate process.
    Returns True if launched, False if already running or failed.
    """
    # Prevent multiple launches (even if user spams Enter)
    if not _try_acquire_lock():
        return False

    try:
        cwd = root_dir or os.getcwd()

        if getattr(sys, "frozen", False):
            # In onefile EXE, run THIS exe with a private flag.
            cmd = [sys.executable, "--menu-one"]
        else:
            # In dev: run as a module
            cmd = [sys.executable, "-m", "menu.Menu_One"]

        # detached Popen so GUI stays alive
        subprocess.Popen(cmd, cwd=cwd)
        return True

    except Exception:
        # If we failed to launch, release lock so user can try again.
        _release_lock()
        return False


def maybe_run_menu_one_from_argv() -> bool:
    """
    Call this near the start of your app entrypoint.
    If '--menu-one' is present, runs Menu One IN THIS PROCESS and exits True.
    Otherwise returns False.
    """
    if "--menu-one" not in sys.argv:
        return False

    # single instance guard in the child process too
    if not _try_acquire_lock():
        return True  # already running elsewhere; silently exit path

    # Import here (safe if bundled)
    import menu.Menu_One as menu_one
    menu_one.main()
    return True
