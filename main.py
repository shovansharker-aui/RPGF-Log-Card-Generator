"""
main.py

Application Entry Point
"""

import ctypes
import sys

WINDOW_TITLE_PREFIX = "RPGF "

_mutex = None


def another_instance_running():
    """Keep a single copy of the app. Returns True (after bringing the
    existing window to the front) when one is already open."""

    try:
        global _mutex

        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

        kernel32.CreateMutexW.restype = ctypes.c_void_p
        kernel32.CreateMutexW.argtypes = [
            ctypes.c_void_p, ctypes.c_bool, ctypes.c_wchar_p]

        # Kept in a module global so the handle lives as long as the app.
        _mutex = kernel32.CreateMutexW(
            None, False, "Renata.RPGF.EngineeringSuite")

        if ctypes.get_last_error() != 183:  # ERROR_ALREADY_EXISTS
            return False

        user32 = ctypes.windll.user32

        found = []

        @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
        def callback(handle, _param):

            length = user32.GetWindowTextLengthW(handle)

            if length and user32.IsWindowVisible(handle):

                buffer = ctypes.create_unicode_buffer(length + 1)

                user32.GetWindowTextW(handle, buffer, length + 1)

                if buffer.value.startswith(WINDOW_TITLE_PREFIX):
                    found.append(handle)

            return True

        user32.EnumWindows(callback, 0)

        if found:
            user32.ShowWindow(found[0], 9)  # SW_RESTORE
            user32.SetForegroundWindow(found[0])

        return True

    except Exception:
        return False


def main():

    if another_instance_running():
        sys.exit(0)

    # Imported here so a second launch exits before loading anything heavy.
    from dashboard import Dashboard

    Dashboard()


if __name__ == "__main__":

    main()
