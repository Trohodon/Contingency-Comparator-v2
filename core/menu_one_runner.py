# core/menu_one_runner.py

from __future__ import annotations

import os
import sys
import tempfile
import traceback
import socket


# Choose a stable port unlikely to collide with common services
# (only used on localhost for single-instance locking)
_LOCK_PORT = 48573


def _is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def _acquire_single_instance_lock() -> socket.socket | None:
    """
    Try to bind a localhost TCP port. If bind fails, another instance is running.
    Returns the bound socket if acquired; caller must keep it alive.
    """
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", _LOCK_PORT))
        s.listen(1)
        return s
    except Exception:
        try:
            s.close()
        except Exception:
            pass
        return None


def _write_fail_log(exc: BaseException) -> str:
    """
    Write a crash log somewhere users can find it.
    Returns the log path.
    """
    try:
        log_path = os.path.join(tempfile.gettempdir(), "dcc_menu_one.log")
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("Menu One launch failed.\n\n")
            f.write("argv:\n")
            f.write(" ".join(sys.argv) + "\n\n")
            f.write("frozen: " + str(_is_frozen()) + "\n\n")
            f.write("traceback:\n")
            f.write("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
        return log_path
    except Exception:
        return ""


def _show_windows_messagebox(title: str, message: str):
    """
    Optional: show a message box without Tk.
    Works only on Windows; safe no-op elsewhere.
    """
    try:
        import ctypes  # type: ignore
        ctypes.windll.user32.MessageBoxW(None, message, title, 0x10)  # MB_ICONERROR
    except Exception:
        pass


def maybe_run_menu_one_from_argv() -> bool:
    """
    If '--menu-one' is in argv, run Menu One and return True.
    If not present, return False.

    IMPORTANT: This must be called very early (before creating Tk).
    """
    if "--menu-one" not in sys.argv:
        return False

    # single-instance guard
    lock = _acquire_single_instance_lock()
    if lock is None:
        # Another Menu One is already running; do nothing (prevents multi-open)
        return True

    try:
        # Import and run the game.
        # This requires menu/ to be packaged (PyInstaller hiddenimports/collect-submodules).
        from menu.Menu_One import main as menu_main  # noqa: F401

        menu_main()
        return True

    except Exception as e:
        log_path = _write_fail_log(e)

        # In a frozen exe, silent failure is annoying—give at least one hint.
        if _is_frozen():
            msg = "Menu One failed to launch."
            if log_path:
                msg += f"\n\nLog written to:\n{log_path}"
            _show_windows_messagebox("Menu One", msg)

        return True

    finally:
        # Keep lock socket alive until we exit this runner
        try:
            lock.close()
        except Exception:
            pass
