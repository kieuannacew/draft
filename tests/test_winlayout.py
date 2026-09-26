"""Giả lập API Windows để kiểm tra logic sắp xếp cửa sổ (máy CI chạy Linux)."""
import ctypes
import os

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
pytest.importorskip("PySide6")
from PySide6.QtWidgets import QApplication, QWidget  # noqa: E402

from mos.qt import winlayout as W  # noqa: E402


class FakeShell32:
    def __init__(self):
        self.calls = []

    def SHAppBarMessage(self, msg, pabd):
        abd = pabd._obj
        self.calls.append((msg, (abd.rc.left, abd.rc.top, abd.rc.right, abd.rc.bottom)))
        if msg == W.ABM_QUERYPOS:          # Taskbar cao 40px ở đáy: Windows đẩy thanh lên trên
            abd.rc.bottom -= 40
        return 1


class FakeUser32:
    def __init__(self, windows):
        self.windows, self.calls = windows, []

    def RegisterWindowMessageW(self, name):
        return 0xC000

    def SetWindowPos(self, *args):
        self.calls.append(("SetWindowPos",) + args)

    def ShowWindow(self, hwnd, cmd):
        self.calls.append(("ShowWindow", hwnd, cmd))

    def IsWindowVisible(self, hwnd):
        return True

    def GetWindowTextLengthW(self, hwnd):
        return len(self.windows[hwnd])

    def GetWindowTextW(self, hwnd, buf, n):
        buf.value = self.windows[hwnd]

    def EnumWindows(self, proc, lparam):
        for hwnd in self.windows:
            proc(hwnd, lparam)


@pytest.fixture
def fake_windows(monkeypatch):
    QApplication.instance() or QApplication([])
    user32 = FakeUser32({101: "DoanhSo.xlsx - Excel", 102: "DoanhSo - Notepad", 103: "BaoCao - Word",
                         104: "doanhso [Protected View] - Excel"})
    shell32 = FakeShell32()
    monkeypatch.setattr(W, "IS_WINDOWS", True)
    monkeypatch.setattr(W, "user32", user32)
    monkeypatch.setattr(W, "shell32", shell32)
    monkeypatch.setattr(W, "_physical_screen", lambda: (0, 0, 1920, 1080, 1.0))
    monkeypatch.setattr(W.atexit, "register", lambda f: None)
    return user32, shell32


def test_dock_reserves_bottom_above_taskbar(fake_windows):
    user32, shell32 = fake_windows
    w = QWidget()
    dock = W.BottomDock(w)
    assert dock.attach(380)
    msgs = [m for m, _ in shell32.calls]
    assert msgs == [W.ABM_NEW, W.ABM_QUERYPOS, W.ABM_SETPOS]
    assert shell32.calls[2][1] == (0, 1040 - 380, 1920, 1040)       # ngay trên Taskbar
    assert user32.calls[-1][2:7] == (W.HWND_TOPMOST, 0, 660, 1920, 380)
    assert dock.top_px == 660
    dock.detach()
    assert shell32.calls[-1][0] == W.ABM_REMOVE and not dock.active


def test_find_and_fit_office_window(fake_windows):
    user32, _ = fake_windows
    assert W.find_office_windows("DoanhSo") == [101, 104]           # bỏ qua Notepad
    dock = W.BottomDock(QWidget())
    dock.attach(380)
    user32.calls.clear()
    W.fit_window(101, dock)
    assert user32.calls == [("ShowWindow", 101, W.SW_RESTORE), ("ShowWindow", 101, W.SW_MAXIMIZE)]
    dock.active = False                                              # không có AppBar → tự đặt kích thước
    user32.calls.clear()
    W.fit_window(101, dock)
    assert user32.calls[-1][:7] == ("SetWindowPos", 101, 0, 0, 0, 1920, 660)


def test_noop_off_windows(monkeypatch):
    QApplication.instance() or QApplication([])
    monkeypatch.setattr(W, "IS_WINDOWS", False)
    assert not W.BottomDock(QWidget()).attach(380)
    assert W.find_office_windows("x") == []


def test_appbardata_layout():
    assert ctypes.sizeof(W.APPBARDATA) >= 36
