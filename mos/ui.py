"""Màu sắc, font và các widget phẳng dùng chung cho mọi màn hình."""
from __future__ import annotations

import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

# ====================================================================== Màu sắc & font

BG = "#f1f5f9"          # nền ứng dụng
CARD = "#ffffff"        # nền thẻ
BORDER = "#e2e8f0"
TEXT = "#0f172a"
MUTED = "#64748b"
PRIMARY = "#2563eb"
PRIMARY_HOVER = "#1d4ed8"
DARK = "#0f172a"
DARK_2 = "#1e293b"
SUCCESS, SUCCESS_BG = "#16a34a", "#dcfce7"
DANGER, DANGER_BG = "#dc2626", "#fee2e2"
WARN, WARN_BG = "#b45309", "#fef3c7"

# Màu nhận diện từng môn: (chữ cái, màu chính, màu nền nhạt)
BRAND = {
    "WORD": ("W", "#2b579a", "#e8eef8"),
    "EXCEL": ("X", "#217346", "#e6f2eb"),
    "POWERPOINT": ("P", "#c43e1c", "#fbebe6"),
}

FAMILY = "Segoe UI"


def F(size=10, weight="normal"):
    return (FAMILY, size, weight)


def _init_theme(root: tk.Tk) -> None:
    global FAMILY
    families = set(tkfont.families(root))
    for fam in ("Segoe UI", "Helvetica Neue", "Inter", "Noto Sans", "DejaVu Sans"):
        if fam in families:
            FAMILY = fam
            break
    root.configure(bg=BG)
    st = ttk.Style(root)
    st.theme_use("clam")
    st.configure("Clean.Treeview", background=CARD, fieldbackground=CARD, foreground=TEXT,
                 rowheight=32, font=F(10), borderwidth=0)
    st.configure("Clean.Treeview.Heading", background=BG, foreground=MUTED, font=F(9, "bold"),
                 relief="flat", padding=(8, 6))
    st.map("Clean.Treeview", background=[("selected", "#dbeafe")], foreground=[("selected", TEXT)])
    st.map("Clean.Treeview.Heading", background=[("active", BORDER)])
    st.configure("Vertical.TScrollbar", background=BORDER, troughcolor=BG, borderwidth=0,
                 arrowcolor=MUTED, relief="flat")


# ====================================================================== Bộ widget


class Btn(tk.Label):
    """Nút phẳng có hiệu ứng hover. kind: primary | secondary | ghost | dark."""

    STYLES = {
        "primary": (PRIMARY, "white", PRIMARY_HOVER),
        "secondary": (CARD, TEXT, "#e2e8f0"),
        "ghost": (CARD, PRIMARY, "#eff6ff"),
        "dark": (DARK_2, "white", "#334155"),
    }

    def __init__(self, parent, text, command, kind="secondary", size=10, padx=16, pady=7, **kw):
        bg, fg, self.hover = self.STYLES[kind]
        border = BORDER if kind == "secondary" else bg
        super().__init__(parent, text=text, bg=bg, fg=fg, font=F(size, "bold"), padx=padx, pady=pady,
                         cursor="hand2", highlightthickness=1, highlightbackground=border, **kw)
        self.base_bg, self.base_fg, self.command, self.enabled = bg, fg, command, True
        self.bind("<Enter>", lambda e: self.enabled and self.config(bg=self.hover))
        self.bind("<Leave>", lambda e: self.config(bg=self.base_bg))
        self.bind("<Button-1>", lambda e: self.enabled and self.command())

    def set_enabled(self, enabled: bool) -> None:
        self.enabled = enabled
        self.config(fg=self.base_fg if enabled else "#94a3b8", cursor="hand2" if enabled else "arrow")


def pill(parent, text, fg, bg, size=9):
    return tk.Label(parent, text=text, fg=fg, bg=bg, font=F(size, "bold"), padx=10, pady=2)


def card(parent, **kw):
    return tk.Frame(parent, bg=CARD, highlightthickness=1, highlightbackground=BORDER, highlightcolor=BORDER,
                    **kw)


def badge(parent, letter, color, size=44, bg=CARD):
    """Ô vuông màu có chữ cái (W / X / P)."""
    c = tk.Canvas(parent, width=size, height=size, bg=bg, highlightthickness=0)
    r = size // 5
    x0, y0, x1, y1 = 1, 1, size - 1, size - 1
    for x, y in ((x0, y0), (x1 - 2 * r, y0), (x0, y1 - 2 * r), (x1 - 2 * r, y1 - 2 * r)):
        c.create_oval(x, y, x + 2 * r, y + 2 * r, fill=color, outline=color)
    c.create_rectangle(x0 + r, y0, x1 - r, y1, fill=color, outline=color)
    c.create_rectangle(x0, y0 + r, x1, y1 - r, fill=color, outline=color)
    c.create_text(size / 2, size / 2, text=letter, fill="white", font=F(int(size / 2.4), "bold"))
    return c


def bind_click(widget, fn):
    widget.bind("<Button-1>", lambda e: fn())
    widget.configure(cursor="hand2")
    for child in widget.winfo_children():
        bind_click(child, fn)


def fmt_time(seconds: int) -> str:
    m, s = divmod(max(int(seconds), 0), 60)
    return f"{m:02d}:{s:02d}"


class Scrollable(tk.Frame):
    """Khung cuộn dọc; nội dung đặt vào `.inner`."""

    def __init__(self, parent, bg=BG):
        super().__init__(parent, bg=bg)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0)
        sb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = tk.Frame(self.canvas, bg=bg)
        win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(win, width=e.width))
        self.canvas.configure(yscrollcommand=sb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.bind_all(seq, self._wheel, add="+")

    def _wheel(self, e):
        if not str(e.widget).startswith(str(self)):
            return
        step = -1 if (getattr(e, "delta", 0) > 0 or e.num == 4) else 1
        self.canvas.yview_scroll(step, "units")


def entry(parent, var, show=None, width=None, **kw):
    """Ô nhập liệu phẳng."""
    e = tk.Entry(parent, textvariable=var, show=show, font=F(10), relief="flat", bg=CARD, fg=TEXT,
                 highlightthickness=1, highlightbackground=BORDER, highlightcolor=PRIMARY,
                 insertbackground=TEXT, **kw)
    if width:
        e.configure(width=width)
    return e


def text_box(parent, height=3):
    return tk.Text(parent, height=height, wrap="word", font=F(10), relief="flat", bg=CARD, fg=TEXT,
                   highlightthickness=1, highlightbackground=BORDER, highlightcolor=PRIMARY,
                   insertbackground=TEXT, padx=6, pady=4)


def label(parent, text, muted=False, bold=False, size=10, bg=None, fg=None, **kw):
    return tk.Label(parent, text=text, bg=bg or parent.cget("bg"), fg=fg or (MUTED if muted else TEXT),
                    font=F(size, "bold" if bold else "normal"), anchor="w", justify="left", **kw)


def make_tree(parent, columns, height=10):
    """Treeview có thanh cuộn. columns: [(id, tiêu đề, độ rộng, co giãn?)]."""
    box = card(parent)
    tree = ttk.Treeview(box, columns=[c[0] for c in columns], show="headings", style="Clean.Treeview",
                        height=height, selectmode="browse")
    for cid, title, width, stretch in columns:
        tree.heading(cid, text=title, anchor="w")
        tree.column(cid, width=width, anchor="w", stretch=stretch)
    tree.tag_configure("ok", foreground=SUCCESS)
    tree.tag_configure("bad", foreground=DANGER)
    tree.tag_configure("muted", foreground=MUTED)
    sb = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    tree.pack(fill="both", expand=True, padx=1, pady=1)
    return box, tree


def modal(win, parent):
    """Biến Toplevel thành hộp thoại: nằm trên cửa sổ cha và chặn thao tác phía sau."""
    win.transient(parent)
    win.update_idletasks()
    x = parent.winfo_rootx() + max(0, (parent.winfo_width() - win.winfo_reqwidth()) // 2)
    y = parent.winfo_rooty() + max(0, (parent.winfo_height() - win.winfo_reqheight()) // 3)
    win.geometry(f"+{x}+{y}")
    try:
        win.grab_set()
    except tk.TclError:
        pass
