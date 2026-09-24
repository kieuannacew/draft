"""Giao diện (Tkinter) mô phỏng GMetrix.

Màn hình:
  * LauncherWindow – chọn đề, chọn chế độ, xem lịch sử.
  * ExamBar       – thanh làm bài nằm ở cạnh dưới màn hình, luôn nổi trên
                    Word/Excel/PowerPoint: yêu cầu, đồng hồ, chuyển dự án.
  * ResultWindow  – điểm (thang 1000, đạt >= 700) và chi tiết đúng/sai.

Chỉ dùng Tkinter có sẵn trong Python; phong cách "phẳng" được tự vẽ bằng
Frame/Label/Canvas (các widget nhỏ ở phần "Bộ widget" bên dưới).
"""
from __future__ import annotations

import time
import tkinter as tk
from tkinter import font as tkfont
from tkinter import messagebox, ttk

from . import core
from .custom import app_dir, load_custom_exams, merge_exams
from .exams import ALL_EXAMS

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
    return tk.Frame(parent, bg=CARD, highlightthickness=1, highlightbackground=BORDER, **kw)


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


# ====================================================================== Launcher


class LauncherWindow:
    def __init__(self, root: tk.Tk):
        self.root = root
        _init_theme(root)
        root.title("Luyện thi MOS")
        custom, errors = load_custom_exams()
        self.exams = merge_exams(ALL_EXAMS, custom)   # đề tự soạn gộp vào môn tương ứng
        root.geometry("900x560")
        root.minsize(820, 600)
        self.exam_idx = 0
        self.mode = "training"
        if errors:
            root.after(300, lambda: messagebox.showwarning(
                "Có đề tự soạn bị lỗi", "Các đề sau chưa được nạp:\n\n• " + "\n• ".join(errors[:15]) +
                "\n\nChạy  py kiem_tra_de.py <thư mục đề>  để xem chi tiết.", parent=root))

        head = tk.Frame(root, bg=DARK, padx=32, pady=22)
        head.pack(fill="x")
        tk.Label(head, text="Luyện thi MOS", bg=DARK, fg="white", font=F(22, "bold")).pack(anchor="w")
        tk.Label(head, text="Làm bài trực tiếp trên Word, Excel, PowerPoint — chấm điểm tự động như GMetrix",
                 bg=DARK, fg="#94a3b8", font=F(11)).pack(anchor="w", pady=(2, 0))

        body = tk.Frame(root, bg=BG, padx=32, pady=20)
        body.pack(fill="both", expand=True)

        # hàng nút pack trước để luôn nằm dưới cùng
        foot = tk.Frame(body, bg=BG)
        foot.pack(fill="x", side="bottom")
        self.last_lbl = tk.Label(foot, bg=BG, fg=MUTED, font=F(10), justify="left")
        self.last_lbl.pack(side="left")
        Btn(foot, "Bắt đầu  →", self.start, kind="primary", size=12, padx=26, pady=10).pack(side="right")
        Btn(foot, "Lịch sử", self.show_history, size=12, pady=10).pack(side="right", padx=10)
        Btn(foot, "Thư mục đề", self.open_exam_dir, size=12, pady=10).pack(side="right")

        self._section(body, "1", "Chọn bài thi")
        exams = tk.Frame(body, bg=BG)
        exams.pack(fill="x", pady=(8, 18))
        self.exam_cards = []
        for col in range(3):
            exams.columnconfigure(col, weight=1, uniform="exam")
        for i, exam in enumerate(self.exams):
            self.exam_cards.append(self._exam_card(exams, i, exam))

        self._section(body, "2", "Chọn chế độ")
        modes = tk.Frame(body, bg=BG)
        modes.pack(fill="x", pady=(8, 18))
        self.mode_cards = {}
        for i, (key, title, desc) in enumerate((
                ("training", "Luyện tập", "Không giới hạn giờ • có gợi ý • kiểm tra từng dự án"),
                ("testing", "Thi thử", "Tính giờ • không gợi ý • chấm điểm khi nộp bài"))):
            modes.columnconfigure(i, weight=1, uniform="mode")
            c = card(modes, padx=16, pady=12)
            c.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 8, 0))
            dot = tk.Canvas(c, width=18, height=18, bg=CARD, highlightthickness=0)
            dot.pack(side="left", padx=(0, 12))
            txt = tk.Frame(c, bg=CARD)
            txt.pack(side="left", fill="x")
            tk.Label(txt, text=title, bg=CARD, fg=TEXT, font=F(12, "bold")).pack(anchor="w")
            tk.Label(txt, text=desc, bg=CARD, fg=MUTED, font=F(9)).pack(anchor="w")
            bind_click(c, lambda k=key: self.select_mode(k))
            self.mode_cards[key] = (c, dot)


        self.select_exam(0)
        self.select_mode("training")
        self.refresh_last()

    @staticmethod
    def _section(parent, num, text):
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x")
        tk.Label(row, text=num, bg=PRIMARY, fg="white", font=F(9, "bold"), width=2).pack(side="left")
        tk.Label(row, text=text, bg=BG, fg=TEXT, font=F(12, "bold")).pack(side="left", padx=8)

    def _exam_card(self, parent, i, exam):
        letter, color, _ = BRAND.get(exam.code, ("?", PRIMARY, "#eff6ff"))
        c = card(parent, padx=14, pady=12)
        row, col = divmod(i, 3)
        c.grid(row=row, column=col, sticky="nsew", padx=(0 if col == 0 else 8, 0), pady=(0 if row == 0 else 8, 0))
        top = tk.Frame(c, bg=CARD)
        top.pack(fill="x")
        badge(top, letter, color, size=38).pack(side="left")
        names = tk.Frame(top, bg=CARD)
        names.pack(side="left", padx=10, fill="x", expand=True)
        short, _, code = exam.name.replace("Microsoft ", "").partition(" (")
        tk.Label(names, text=short, bg=CARD, fg=TEXT, font=F(13 if len(short) < 16 else 10, "bold"),
                 wraplength=170, justify="left", anchor="w").pack(anchor="w")
        tk.Label(names, text=code.rstrip(")"), bg=CARD, fg=MUTED, font=F(9)).pack(anchor="w")
        n_tasks = sum(len(p.tasks) for p in exam.projects)
        tk.Label(c, text=f"{len(exam.projects)} dự án  •  {n_tasks} nhiệm vụ  •  {exam.minutes} phút",
                 bg=CARD, fg=MUTED, font=F(9)).pack(anchor="w", pady=(8, 0))
        best = max((h["score"] for h in core.load_history() if h["exam"] == exam.name), default=None)
        tk.Label(c, text=f"Điểm cao nhất: {best}" if best is not None else "Chưa làm lần nào",
                 bg=CARD, fg=color if best is not None else MUTED, font=F(9, "bold")).pack(anchor="w")
        bind_click(c, lambda: self.select_exam(i))
        return c, color

    def select_exam(self, idx):
        self.exam_idx = idx
        for i, (c, color) in enumerate(self.exam_cards):
            c.config(highlightbackground=color if i == idx else BORDER, highlightthickness=2 if i == idx else 1)

    def select_mode(self, key):
        self.mode = key
        for k, (c, dot) in self.mode_cards.items():
            on = k == key
            c.config(highlightbackground=PRIMARY if on else BORDER, highlightthickness=2 if on else 1)
            dot.delete("all")
            dot.create_oval(2, 2, 16, 16, outline=PRIMARY if on else "#cbd5e1", width=2)
            if on:
                dot.create_oval(6, 6, 12, 12, fill=PRIMARY, outline=PRIMARY)

    def refresh_last(self):
        history = core.load_history()
        if history:
            h = history[-1]
            name = h["exam"].replace("Microsoft ", "").split(" (")[0]
            self.last_lbl.config(text=f"Lần gần nhất: {name} — {h['score']}/1000 "
                                      f"({'Đạt' if h['passed'] else 'Chưa đạt'})")
        else:
            self.last_lbl.config(text="Mẹo: làm xong mỗi dự án nhớ bấm Ctrl+S để lưu file.")

    def start(self) -> None:
        start_exam(self.root, self.exams[self.exam_idx], self.mode)

    def open_exam_dir(self) -> None:
        folder = app_dir() / "de_thi"
        folder.mkdir(exist_ok=True)
        try:
            core.open_in_office(folder)
        except OSError as exc:
            messagebox.showerror("Lỗi", str(exc))

    def show_history(self) -> None:
        HistoryWindow(self.root)


def start_exam(root, exam, mode) -> None:
    try:
        session = core.new_session(exam, mode)
    except Exception:
        messagebox.showerror("Lỗi", "Không tạo được file bài thi:\n" + core.format_exception())
        return
    root.withdraw()
    ExamBar(root, session)


# ====================================================================== History


class HistoryWindow:
    def __init__(self, root):
        w = tk.Toplevel(root, bg=BG)
        w.title("Lịch sử làm bài")
        w.geometry("820x460")
        history = core.load_history()

        stats = tk.Frame(w, bg=BG, padx=20, pady=16)
        stats.pack(fill="x")
        passed = sum(h["passed"] for h in history)
        for i, (label, value) in enumerate((
                ("Số lần làm", str(len(history))),
                ("Điểm cao nhất", str(max((h["score"] for h in history), default="—"))),
                ("Tỉ lệ đạt", f"{round(100 * passed / len(history))}%" if history else "—"))):
            stats.columnconfigure(i, weight=1, uniform="s")
            c = card(stats, padx=16, pady=10)
            c.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 8, 0))
            tk.Label(c, text=label, bg=CARD, fg=MUTED, font=F(9)).pack(anchor="w")
            tk.Label(c, text=value, bg=CARD, fg=TEXT, font=F(18, "bold")).pack(anchor="w")

        box = card(w)
        box.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        cols = ("time", "exam", "mode", "duration", "score", "result")
        tree = ttk.Treeview(box, columns=cols, show="headings", style="Clean.Treeview")
        for c, text, width in zip(cols, ("Thời gian", "Bài thi", "Chế độ", "Làm trong", "Điểm", "Kết quả"),
                                  (130, 230, 90, 90, 90, 90)):
            tree.heading(c, text=text, anchor="w")
            tree.column(c, width=width, anchor="w")
        tree.tag_configure("ok", foreground=SUCCESS)
        tree.tag_configure("bad", foreground=DANGER)
        for h in reversed(history):
            tree.insert("", "end", tags=("ok" if h["passed"] else "bad",), values=(
                h["time"], h["exam"], "Luyện tập" if h["mode"] == "training" else "Thi thử",
                fmt_time(h["duration"]) if "duration" in h else "—",
                f'{h["score"]}/1000', "Đạt" if h["passed"] else "Chưa đạt"))
        tree.pack(fill="both", expand=True, padx=1, pady=1)
        if not history:
            tk.Label(box, text="Chưa có lần làm bài nào.", bg=CARD, fg=MUTED, font=F(10)).place(
                relx=0.5, rely=0.5, anchor="center")


# ====================================================================== Exam bar


class ExamBar:
    HEIGHT = 380

    def __init__(self, root: tk.Tk, session: core.Session):
        self.root = root
        self.s = session
        self.training = session.mode == "training"
        self.idx = 0
        self.started = time.time()
        self.seconds = 0 if self.training else session.exam.minutes * 60
        self.finished = False
        self.checked: dict[tuple, bool] = {}     # kết quả "Kiểm tra dự án" gần nhất
        self.open_hints: set = set()
        letter, self.color, self.tint = BRAND.get(session.exam.code, ("?", PRIMARY, "#eff6ff"))

        w = self.win = tk.Toplevel(root, bg=BG)
        w.title(f"MOS – {session.exam.name}")
        w.attributes("-topmost", True)
        sw, sh = w.winfo_screenwidth(), w.winfo_screenheight()
        w.geometry(f"{sw}x{self.HEIGHT}+0+{max(0, sh - self.HEIGHT - 60)}")
        w.protocol("WM_DELETE_WINDOW", self.ask_quit)

        # --- thanh trên (tối): môn thi, tab dự án, đồng hồ
        top = tk.Frame(w, bg=DARK, padx=14, pady=8)
        top.pack(fill="x")
        badge(top, letter, self.color, size=30, bg=DARK).pack(side="left")
        tk.Label(top, text=session.exam.name, bg=DARK, fg="white", font=F(11, "bold")).pack(side="left", padx=10)
        pill(top, "LUYỆN TẬP" if self.training else "THI THỬ",
             "white", PRIMARY if self.training else "#7c3aed").pack(side="left")
        self.tabs_frame = tk.Frame(top, bg=DARK)
        self.tabs_frame.pack(side="left", padx=24)
        self.timer_lbl = tk.Label(top, bg=DARK_2, fg="white", font=("Consolas", 15, "bold"), padx=14, pady=2)
        self.timer_lbl.pack(side="right")
        tk.Label(top, text="⏱", bg=DARK, fg="#94a3b8", font=F(13)).pack(side="right", padx=6)

        # --- thanh dưới: nút thao tác (pack trước phần thân để không bị che)
        bottom = tk.Frame(w, bg=CARD, padx=14, pady=8, highlightthickness=1, highlightbackground=BORDER)
        bottom.pack(side="bottom", fill="x")
        Btn(bottom, "Mở file", self.open_file).pack(side="left")
        Btn(bottom, "Làm lại dự án", self.reset_project).pack(side="left", padx=6)
        if self.training:
            Btn(bottom, "✓  Kiểm tra dự án", self.check_current, kind="ghost").pack(side="left")
        self.progress_lbl = tk.Label(bottom, bg=CARD, fg=MUTED, font=F(10))
        self.progress_lbl.pack(side="left", padx=18)
        Btn(bottom, "Nộp bài", self.submit, kind="primary", padx=22).pack(side="right")
        self.next_btn = Btn(bottom, "Dự án sau  ›", lambda: self.go(self.idx + 1))
        self.next_btn.pack(side="right", padx=6)
        self.prev_btn = Btn(bottom, "‹  Dự án trước", lambda: self.go(self.idx - 1))
        self.prev_btn.pack(side="right")

        # --- mô tả dự án
        intro = tk.Frame(w, bg=self.tint, padx=16, pady=7)
        intro.pack(fill="x")
        self.intro_lbl = tk.Label(intro, bg=self.tint, fg=TEXT, font=F(10), anchor="w", justify="left")
        self.intro_lbl.pack(side="left")
        self.file_lbl = tk.Label(intro, bg=self.tint, fg=self.color, font=F(9, "bold"))
        self.file_lbl.pack(side="right")

        # --- danh sách nhiệm vụ
        self.list = Scrollable(w)
        self.list.pack(fill="both", expand=True, padx=12, pady=8)

        self.show_project()
        self.open_file()
        self.tick()

    # ------------------------------------------------------------ hiển thị
    def show_project(self) -> None:
        project = self.s.exam.projects[self.idx]
        n = len(self.s.exam.projects)

        for child in self.tabs_frame.winfo_children():
            child.destroy()
        for i, p in enumerate(self.s.exam.projects):
            on = i == self.idx
            short = p.name.split("–")[-1].strip()
            if n > 4:
                short = short[:12] + ("…" if len(short) > 12 else "")
            tab = tk.Label(self.tabs_frame, text=f"{i + 1}  {short}", font=F(9, "bold"), padx=12, pady=3,
                           bg=CARD if on else DARK_2, fg=TEXT if on else "#cbd5e1", cursor="hand2")
            tab.pack(side="left", padx=3)
            tab.bind("<Button-1>", lambda e, k=i: self.go(k))

        self.intro_lbl.config(text=f"{project.intro}   Làm xong nhớ lưu file (Ctrl+S).")
        self.file_lbl.config(text=f"📄 {project.filename}")
        self.prev_btn.set_enabled(self.idx > 0)
        self.next_btn.set_enabled(self.idx < n - 1)

        for child in self.list.inner.winfo_children():
            child.destroy()
        wrap = max(400, self.win.winfo_screenwidth() - 520)
        for t_idx, task in enumerate(project.tasks):
            self._task_row(t_idx, task, wrap)
        self.list.canvas.yview_moveto(0)
        self.update_progress()

    def _task_row(self, t_idx, task, wrap):
        key = (self.idx, t_idx)
        row = card(self.list.inner, padx=12, pady=8)
        row.pack(fill="x", pady=3)
        line = tk.Frame(row, bg=CARD)
        line.pack(fill="x")

        num = tk.Canvas(line, width=28, height=28, bg=CARD, highlightthickness=0)
        done = key in self.s.done
        num.create_oval(2, 2, 26, 26, fill=self.color if done else CARD, outline=self.color, width=2)
        num.create_text(14, 14, text=str(t_idx + 1), fill="white" if done else self.color, font=F(10, "bold"))
        num.pack(side="left")

        tk.Label(line, text=task.title, bg=CARD, fg=TEXT, font=F(10), wraplength=wrap, justify="left",
                 anchor="w").pack(side="left", padx=12, fill="x", expand=True)

        right = tk.Frame(line, bg=CARD)
        right.pack(side="right")
        if key in self.checked:
            ok = self.checked[key]
            pill(right, "✓ Đúng" if ok else "✗ Sai", SUCCESS if ok else DANGER,
                 SUCCESS_BG if ok else DANGER_BG).pack(side="left", padx=4)
        if self.training:
            hint_on = key in self.open_hints
            h = tk.Label(right, text="Ẩn gợi ý" if hint_on else "💡 Gợi ý", bg=CARD, fg=WARN,
                         font=F(9, "bold"), cursor="hand2", padx=6)
            h.pack(side="left")
            h.bind("<Button-1>", lambda e: self._toggle(self.open_hints, key))
        marked = key in self.s.marked
        m = tk.Label(right, text="⚑ Đã đánh dấu" if marked else "⚐ Đánh dấu", font=F(9, "bold"),
                     bg=WARN_BG if marked else CARD, fg=WARN if marked else MUTED, cursor="hand2", padx=6)
        m.pack(side="left", padx=4)
        m.bind("<Button-1>", lambda e: self._toggle(self.s.marked, key))
        d = tk.Label(right, text="☑ Đã làm" if done else "☐ Đã làm", font=F(9, "bold"),
                     bg=SUCCESS_BG if done else CARD, fg=SUCCESS if done else MUTED, cursor="hand2", padx=6)
        d.pack(side="left")
        d.bind("<Button-1>", lambda e: self._toggle(self.s.done, key))

        if self.training and key in self.open_hints:
            tk.Label(row, text=f"Cách làm: {task.hint}", bg=WARN_BG, fg="#78350f", font=F(9),
                     wraplength=wrap + 200, justify="left", anchor="w", padx=10, pady=6
                     ).pack(fill="x", pady=(8, 0), padx=(40, 0))

    def _toggle(self, store: set, key) -> None:
        (store.discard if key in store else store.add)(key)
        y = self.list.canvas.yview()[0]
        self.show_project()
        self.list.canvas.yview_moveto(y)

    def update_progress(self) -> None:
        total = sum(len(p.tasks) for p in self.s.exam.projects)
        text = f"Đã làm {len(self.s.done)}/{total}"
        if self.s.marked:
            text += f"   •   Đánh dấu {len(self.s.marked)}"
        self.progress_lbl.config(text=text)

    def tick(self) -> None:
        if self.finished:
            return
        if self.training:
            self.seconds += 1
        else:
            self.seconds -= 1
            if self.seconds <= 0:
                self.timer_lbl.config(text="00:00")
                messagebox.showwarning("Hết giờ", "Đã hết thời gian. Bài sẽ được nộp tự động.\n"
                                       "(Các thay đổi chưa lưu sẽ không được tính.)", parent=self.win)
                self.finish()
                return
        urgent = not self.training and self.seconds < 300
        self.timer_lbl.config(text=fmt_time(self.seconds), bg=DANGER if urgent else DARK_2)
        self.win.after(1000, self.tick)

    # ------------------------------------------------------------ thao tác
    def open_file(self) -> None:
        path = self.s.file_of(self.idx)
        try:
            core.open_in_office(path)
        except OSError as exc:
            messagebox.showerror("Không mở được file",
                                 f"{exc}\n\nHãy tự mở file:\n{path}", parent=self.win)

    def go(self, idx: int) -> None:
        if 0 <= idx < len(self.s.exam.projects) and idx != self.idx:
            self.idx = idx
            self.show_project()
            self.open_file()

    def reset_project(self) -> None:
        path = self.s.file_of(self.idx)
        if core.file_is_open(path):
            messagebox.showwarning("File đang mở", "Hãy ĐÓNG file trong Office trước khi làm lại.",
                                   parent=self.win)
            return
        if messagebox.askyesno("Làm lại", "Tạo lại file gốc cho dự án này?\n(Bài cũ được lưu với tên *_cu)",
                               parent=self.win):
            core.reset_project(self.s, self.idx)
            for t in range(len(self.s.exam.projects[self.idx].tasks)):
                self.checked.pop((self.idx, t), None)
            self.show_project()
            self.open_file()

    def _confirm_saved(self, path) -> bool:
        if core.file_is_open(path):
            return messagebox.askyesno(
                "File đang mở",
                "File vẫn đang mở trong Office. Chỉ những gì đã LƯU mới được chấm.\n\n"
                "Bạn đã bấm Ctrl+S chưa? Chọn Yes để chấm.", parent=self.win)
        return True

    def check_current(self) -> None:
        if not self._confirm_saved(self.s.file_of(self.idx)):
            return
        results = core.check_project(self.s, self.idx)
        for t_idx, r in enumerate(results):
            self.checked[(self.idx, t_idx)] = r.correct
        self.show_project()
        errs = [r.error for r in results if r.error]
        if errs:
            messagebox.showwarning("Không đọc được file", f"Lỗi: {errs[0]}\n\nHãy lưu lại file rồi thử lại.",
                                   parent=self.win)

    def submit(self) -> None:
        opened = [i for i in range(len(self.s.exam.projects)) if core.file_is_open(self.s.file_of(i))]
        msg = "Nộp bài và chấm điểm?"
        if opened:
            msg = "Vẫn còn file đang mở trong Office. Chỉ phần đã LƯU mới được chấm.\n\n" + msg
        if self.s.marked:
            msg = f"Bạn còn {len(self.s.marked)} nhiệm vụ đánh dấu xem lại.\n" + msg
        if messagebox.askyesno("Nộp bài", msg, parent=self.win):
            self.finish()

    def finish(self) -> None:
        self.finished = True
        report = core.grade(self.s)
        report["duration"] = int(time.time() - self.started)
        try:
            core.save_history(report)
        except OSError:
            pass
        self.win.destroy()
        ResultWindow(self.root, self.s, report)

    def ask_quit(self) -> None:
        if messagebox.askyesno("Thoát", "Thoát bài thi? Kết quả sẽ không được chấm.", parent=self.win):
            self.finished = True
            self.win.destroy()
            self.root.deiconify()


# ====================================================================== Result


class ResultWindow:
    def __init__(self, root: tk.Tk, session: core.Session, report: dict):
        self.root = root
        self.session = session
        w = self.win = tk.Toplevel(root, bg=BG)
        w.title("Kết quả")
        w.geometry("980x700")
        w.minsize(900, 620)
        w.protocol("WM_DELETE_WINDOW", self.close)
        passed = report["passed"]
        color = SUCCESS if passed else DANGER

        # --- vòng điểm + thống kê
        head = card(w, padx=24, pady=20)
        head.pack(fill="x", padx=20, pady=(20, 10))
        ring = tk.Canvas(head, width=150, height=150, bg=CARD, highlightthickness=0)
        ring.pack(side="left")
        ring.create_oval(10, 10, 140, 140, outline=BORDER, width=14)
        extent = -359.9 * report["score"] / core.MAX_SCORE
        if extent:
            ring.create_arc(10, 10, 140, 140, start=90, extent=extent, style="arc", outline=color, width=14)
        ring.create_text(75, 68, text=str(report["score"]), fill=TEXT,
                         font=F(24 if report["score"] >= 1000 else 28, "bold"))
        ring.create_text(75, 98, text="/ 1000", fill=MUTED, font=F(10))

        info = tk.Frame(head, bg=CARD)
        info.pack(side="left", padx=24, fill="both", expand=True)
        pill(info, "ĐẠT" if passed else "CHƯA ĐẠT", color, SUCCESS_BG if passed else DANGER_BG,
             size=11).pack(anchor="w")
        tk.Label(info, text=session.exam.name, bg=CARD, fg=TEXT, font=F(16, "bold")).pack(anchor="w", pady=(8, 0))
        tk.Label(info, text=("Chúc mừng! Bạn đã vượt qua điểm đạt." if passed else
                             f"Cần thêm {core.PASS_SCORE - report['score']} điểm để đạt. Xem lại các nhiệm vụ sai bên dưới."),
                 bg=CARD, fg=MUTED, font=F(10)).pack(anchor="w")

        stats = tk.Frame(info, bg=CARD)
        stats.pack(anchor="w", pady=(12, 0))
        for label, value in (("Đúng", f'{report["correct"]}/{report["total"]}'),
                             ("Thời gian", fmt_time(report.get("duration", 0))),
                             ("Điểm đạt", str(core.PASS_SCORE)),
                             ("Chế độ", "Luyện tập" if session.mode == "training" else "Thi thử")):
            box = tk.Frame(stats, bg=BG, padx=14, pady=6)
            box.pack(side="left", padx=(0, 8))
            tk.Label(box, text=label, bg=BG, fg=MUTED, font=F(9)).pack(anchor="w")
            tk.Label(box, text=value, bg=BG, fg=TEXT, font=F(13, "bold")).pack(anchor="w")

        # --- điểm theo từng dự án
        per = tk.Frame(w, bg=BG)
        per.pack(fill="x", padx=20)
        self.results = report["results"]
        for i, project in enumerate(session.exam.projects):
            rs = [r for r in self.results if r.project == project.name]
            ok = sum(r.correct for r in rs)
            per.columnconfigure(i, weight=1, uniform="p")
            c = card(per, padx=14, pady=10)
            c.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 8, 0))
            top = tk.Frame(c, bg=CARD)
            top.pack(fill="x")
            tk.Label(top, text=project.name, bg=CARD, fg=TEXT, font=F(10, "bold")).pack(side="left")
            tk.Label(top, text=f"{ok}/{len(rs)}", bg=CARD, fg=MUTED, font=F(10, "bold")).pack(side="right")
            bar = tk.Canvas(c, height=8, bg=BORDER, highlightthickness=0)
            bar.pack(fill="x", pady=(8, 0))
            ratio = ok / len(rs) if rs else 0
            bar.bind("<Configure>", lambda e, b=bar, r=ratio: (
                b.delete("all"), b.create_rectangle(0, 0, e.width * r, 8, fill=SUCCESS if r >= .7 else WARN,
                                                    outline="")))

        btns = tk.Frame(w, bg=BG, padx=20, pady=14)
        btns.pack(side="bottom", fill="x")
        Btn(btns, "Mở thư mục bài làm", lambda: core.open_in_office(session.workdir)).pack(side="left")
        Btn(btns, "Về màn hình chính", self.close, kind="primary").pack(side="right")
        Btn(btns, "Làm lại đề này", self.retry).pack(side="right", padx=8)

        self.hint = tk.Label(w, bg=WARN_BG, fg="#78350f", font=F(10), anchor="w", justify="left",
                             padx=14, pady=10, wraplength=920,
                             text="💡 Chọn một nhiệm vụ ở bảng trên để xem cách làm.")
        self.hint.pack(side="bottom", fill="x", padx=20)

        # --- danh sách nhiệm vụ
        box = card(w)
        box.pack(fill="both", expand=True, padx=20, pady=10)
        cols = ("result", "project", "task")
        tree = self.tree = ttk.Treeview(box, columns=cols, show="headings", style="Clean.Treeview", height=6)
        for c, text, width, stretch in (("result", "", 80, False), ("project", "Dự án", 210, False),
                                        ("task", "Nhiệm vụ", 600, True)):
            tree.heading(c, text=text, anchor="w")
            tree.column(c, width=width, anchor="w", stretch=stretch)
        tree.tag_configure("ok", foreground=SUCCESS)
        tree.tag_configure("bad", foreground=DANGER)
        for i, r in enumerate(self.results):
            tree.insert("", "end", iid=str(i), tags=("ok" if r.correct else "bad",),
                        values=("✓ Đúng" if r.correct else "✗ Sai", r.project.split("–")[-1].strip(), r.task))
        sb = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        tree.pack(fill="both", expand=True, padx=1, pady=1)
        tree.bind("<<TreeviewSelect>>", self.on_select)


    def on_select(self, _event) -> None:
        sel = self.tree.selection()
        if sel:
            r = self.results[int(sel[0])]
            text = f"💡 Cách làm: {r.hint}"
            if r.error:
                text += f"\n(Lỗi khi đọc file: {r.error})"
            self.hint.config(text=text)

    def retry(self) -> None:
        self.win.destroy()
        start_exam(self.root, self.session.exam, self.session.mode)

    def close(self) -> None:
        self.win.destroy()
        for child in self.root.winfo_children():
            child.destroy()
        LauncherWindow(self.root)   # dựng lại để cập nhật "điểm cao nhất"
        self.root.deiconify()


def main() -> None:
    root = tk.Tk()
    LauncherWindow(root)
    root.mainloop()
