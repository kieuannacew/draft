"""Sắp xếp cửa sổ trên Windows khi làm bài.

* BottomDock: đăng ký thanh làm bài là một "AppBar" ở đáy màn hình (giống Taskbar).
  Windows sẽ trừ phần này ra khỏi vùng làm việc, nên cửa sổ Word/Excel/PowerPoint
  khi phóng to (Maximize) tự vừa khít phần trống phía trên thanh đề.
* OfficeFitter: sau khi mở file, dò cửa sổ Office có tên file đó rồi phóng to nó
  vào phần trống (nếu không đăng ký được AppBar thì đặt kích thước trực tiếp).

Trên hệ điều hành khác Windows, mọi hàm đều không làm gì.
"""
from __future__ import annotations

import atexit
import ctypes
import sys

from PySide6.QtCore import QObject, QTimer
from PySide6.QtGui import QGuiApplication

IS_WINDOWS = sys.platform == "win32"
OFFICE_APPS = ("Excel", "Word", "PowerPoint")

from ctypes import wintypes  # noqa: E402  (import được trên mọi hệ điều hành)


class APPBARDATA(ctypes.Structure):
    _fields_ = [("cbSize", wintypes.DWORD), ("hWnd", wintypes.HWND), ("uCallbackMessage", wintypes.UINT),
                ("uEdge", wintypes.UINT), ("rc", wintypes.RECT), ("lParam", wintypes.LPARAM)]


EnumWindowsProc = getattr(ctypes, "WINFUNCTYPE", ctypes.CFUNCTYPE)(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
user32 = shell32 = None
if IS_WINDOWS:  # pragma: no cover - chỉ chạy trên Windows
    user32 = ctypes.windll.user32
    shell32 = ctypes.windll.shell32
    shell32.SHAppBarMessage.argtypes = [wintypes.DWORD, ctypes.POINTER(APPBARDATA)]
    shell32.SHAppBarMessage.restype = ctypes.c_size_t
    user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                    ctypes.c_int, wintypes.UINT]
    user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
    user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
    user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user32.IsWindowVisible.argtypes = [wintypes.HWND]

ABM_NEW, ABM_REMOVE, ABM_QUERYPOS, ABM_SETPOS = 0, 1, 2, 3
ABE_BOTTOM = 3
HWND_TOPMOST = -1
SWP_NOZORDER, SWP_NOACTIVATE = 0x0004, 0x0010
SW_MAXIMIZE, SW_RESTORE = 3, 9


def _physical_screen():
    """(trái, trên, phải, dưới, tỉ lệ DPI) của màn hình chính, theo pixel thật."""
    screen = QGuiApplication.primaryScreen()
    dpr = screen.devicePixelRatio()
    g = screen.geometry()
    return (round(g.x() * dpr), round(g.y() * dpr), round((g.x() + g.width()) * dpr),
            round((g.y() + g.height()) * dpr), dpr)


class BottomDock:
    """Giữ chỗ ở đáy màn hình cho thanh làm bài (Windows AppBar)."""

    def __init__(self, widget):
        self.widget = widget
        self.active = False
        self.abd = None
        self.top_px = None           # mép trên của thanh (pixel thật) – dùng khi phải tự đặt cửa sổ Office

    def attach(self, height: int) -> bool:
        if not IS_WINDOWS:
            return False
        try:  # pragma: no cover - chỉ chạy trên Windows
            left, _, right, bottom, dpr = _physical_screen()
            h = round(height * dpr)
            abd = APPBARDATA()
            abd.cbSize = ctypes.sizeof(APPBARDATA)
            abd.hWnd = int(self.widget.winId())
            abd.uCallbackMessage = user32.RegisterWindowMessageW("MOS_PRACTICE_APPBAR")
            if not shell32.SHAppBarMessage(ABM_NEW, ctypes.byref(abd)):
                return False
            self.abd, self.active = abd, True
            atexit.register(self.detach)
            abd.uEdge = ABE_BOTTOM
            abd.rc = wintypes.RECT(left, bottom - h, right, bottom)
            shell32.SHAppBarMessage(ABM_QUERYPOS, ctypes.byref(abd))     # tránh Taskbar
            abd.rc.top = abd.rc.bottom - h
            shell32.SHAppBarMessage(ABM_SETPOS, ctypes.byref(abd))
            rc = abd.rc
            user32.SetWindowPos(abd.hWnd, HWND_TOPMOST, rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top,
                                SWP_NOACTIVATE)
            self.top_px = rc.top
            return True
        except Exception:
            self.detach()
            return False

    def detach(self) -> None:
        if IS_WINDOWS and self.active and self.abd is not None:
            try:  # pragma: no cover - chỉ chạy trên Windows
                shell32.SHAppBarMessage(ABM_REMOVE, ctypes.byref(self.abd))
            except Exception:
                pass
        self.active = False


def find_office_windows(stem: str) -> list[int]:
    """Cửa sổ Word/Excel/PowerPoint đang hiện có tiêu đề chứa tên file (không đuôi)."""
    if not IS_WINDOWS:
        return []
    found: list[int] = []
    target = stem.casefold()

    def callback(hwnd, _):  # pragma: no cover - chỉ chạy trên Windows
        if user32.IsWindowVisible(hwnd):
            n = user32.GetWindowTextLengthW(hwnd)
            if n:
                buf = ctypes.create_unicode_buffer(n + 1)
                user32.GetWindowTextW(hwnd, buf, n + 1)
                title = buf.value
                if target in title.casefold() and any(app in title for app in OFFICE_APPS):
                    found.append(hwnd)
        return True

    user32.EnumWindows(EnumWindowsProc(callback), 0)
    return found


def fit_window(hwnd: int, dock: BottomDock) -> None:
    """Đưa cửa sổ Office vào phần trống phía trên thanh làm bài."""
    if not IS_WINDOWS:
        return
    try:  # pragma: no cover - chỉ chạy trên Windows
        if dock.active:
            user32.ShowWindow(hwnd, SW_RESTORE)
            user32.ShowWindow(hwnd, SW_MAXIMIZE)     # vùng làm việc đã trừ thanh đề → vừa khít
        else:
            left, top, right, _, _ = _physical_screen()
            bottom = dock.top_px or round(dock.widget.frameGeometry().top() *
                                          QGuiApplication.primaryScreen().devicePixelRatio())
            user32.ShowWindow(hwnd, SW_RESTORE)
            user32.SetWindowPos(hwnd, 0, left, top, right - left, bottom - top, SWP_NOZORDER | SWP_NOACTIVATE)
    except Exception:
        pass


class OfficeFitter(QObject):
    """Sau khi mở file: dò cửa sổ Office trong ~30 giây rồi thu nhỏ cho vừa phần trống."""

    def __init__(self, dock: BottomDock, parent=None):
        super().__init__(parent)
        self.dock = dock
        self.stem = ""
        self.tries = 0
        self.timer = QTimer(self)
        self.timer.setInterval(500)
        self.timer.timeout.connect(self._tick)

    def watch(self, stem: str) -> None:
        if not IS_WINDOWS:
            return
        self.stem, self.tries = stem, 0
        self.timer.start()

    def _tick(self):  # pragma: no cover - chỉ chạy trên Windows
        self.tries += 1
        hwnds = find_office_windows(self.stem)
        if hwnds:
            QTimer.singleShot(700, lambda: [fit_window(h, self.dock) for h in find_office_windows(self.stem)])
            self.timer.stop()
        elif self.tries > 60:
            self.timer.stop()

