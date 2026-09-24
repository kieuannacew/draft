"""Giao diện (Tkinter) mô phỏng GMetrix.

Màn hình:
  * LauncherWindow – chọn đề, chọn chế độ, xem lịch sử.
  * ExamBar       – thanh làm bài nằm ở cạnh dưới màn hình, luôn nổi trên
                    Word/Excel/PowerPoint: yêu cầu, đồng hồ, chuyển dự án.
  * ResultWindow  – điểm (thang 1000, đạt >= 700) và chi tiết đúng/sai.
"""
from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from . import core
from .exams import ALL_EXAMS

PRIMARY = "#1f5fa8"
OK_COLOR = "#1a7f37"
BAD_COLOR = "#c62828"
FONT = ("Segoe UI", 10)
FONT_B = ("Segoe UI", 10, "bold")


def _style(root: tk.Tk) -> None:
    st = ttk.Style(root)
    if "vista" in st.theme_names():
        st.theme_use("vista")
    elif "clam" in st.theme_names():
        st.theme_use("clam")
    st.configure("Header.TLabel", font=("Segoe UI", 16, "bold"), foreground=PRIMARY)
    st.configure("Timer.TLabel", font=("Consolas", 16, "bold"))
    st.configure("Accent.TButton", font=FONT_B)


# ====================================================================== Launcher


class LauncherWindow:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("Luyện thi MOS")
        root.geometry("660x430")
        _style(root)

        frm = ttk.Frame(root, padding=20)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Luyện thi MOS", style="Header.TLabel").pack(anchor="w")
        ttk.Label(frm, text="Làm bài trực tiếp trên Word, Excel, PowerPoint – chấm điểm tự động.",
                  font=FONT).pack(anchor="w", pady=(0, 14))

        box = ttk.LabelFrame(frm, text=" Chọn bài thi ", padding=10)
        box.pack(fill="x")
        self.exam_var = tk.IntVar(value=0)
        for i, exam in enumerate(ALL_EXAMS):
            n_tasks = sum(len(p.tasks) for p in exam.projects)
            ttk.Radiobutton(box, variable=self.exam_var, value=i,
                            text=f"{exam.name}  –  {len(exam.projects)} dự án, {n_tasks} nhiệm vụ"
                            ).pack(anchor="w", pady=2)

        mode = ttk.LabelFrame(frm, text=" Chế độ ", padding=10)
        mode.pack(fill="x", pady=12)
        self.mode_var = tk.StringVar(value="training")
        ttk.Radiobutton(mode, variable=self.mode_var, value="training",
                        text="Luyện tập (Training) – không giới hạn giờ, có gợi ý, kiểm tra từng dự án"
                        ).pack(anchor="w", pady=2)
        ttk.Radiobutton(mode, variable=self.mode_var, value="testing",
                        text="Thi thử (Testing) – 50 phút, không gợi ý, chấm khi nộp bài"
                        ).pack(anchor="w", pady=2)

        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(8, 0))
        ttk.Button(btns, text="Bắt đầu", style="Accent.TButton", command=self.start).pack(side="left")
        ttk.Button(btns, text="Lịch sử", command=self.show_history).pack(side="left", padx=8)
        ttk.Button(btns, text="Thoát", command=root.destroy).pack(side="right")

    def start(self) -> None:
        exam = ALL_EXAMS[self.exam_var.get()]
        try:
            session = core.new_session(exam, self.mode_var.get())
        except Exception:
            messagebox.showerror("Lỗi", "Không tạo được file bài thi:\n" + core.format_exception())
            return
        self.root.withdraw()
        ExamBar(self.root, session)

    def show_history(self) -> None:
        win = tk.Toplevel(self.root)
        win.title("Lịch sử làm bài")
        win.geometry("720x360")
        cols = ("time", "exam", "mode", "score", "result")
        tree = ttk.Treeview(win, columns=cols, show="headings")
        for c, text, w in zip(cols, ("Thời gian", "Bài thi", "Chế độ", "Điểm", "Kết quả"),
                              (130, 240, 90, 80, 90)):
            tree.heading(c, text=text)
            tree.column(c, width=w, anchor="w")
        tree.pack(fill="both", expand=True, padx=10, pady=10)
        for h in reversed(core.load_history()):
            tree.insert("", "end", values=(
                h["time"], h["exam"], "Luyện tập" if h["mode"] == "training" else "Thi thử",
                f'{h["score"]}/1000', "ĐẠT" if h["passed"] else "Chưa đạt"))


# ====================================================================== Exam bar


class ExamBar:
    HEIGHT = 340

    def __init__(self, root: tk.Tk, session: core.Session):
        self.root = root
        self.s = session
        self.training = session.mode == "training"
        self.idx = 0
        self.seconds = 0 if self.training else session.exam.minutes * 60
        self.finished = False
        self.task_vars: list[tuple[tk.BooleanVar, tk.BooleanVar]] = []

        w = self.win = tk.Toplevel(root)
        w.title(f"MOS – {session.exam.name}")
        w.attributes("-topmost", True)
        sw, sh = w.winfo_screenwidth(), w.winfo_screenheight()
        w.geometry(f"{sw}x{self.HEIGHT}+0+{max(0, sh - self.HEIGHT - 60)}")
        w.protocol("WM_DELETE_WINDOW", self.ask_quit)

        # --- hàng trên: tên bài thi, dự án, đồng hồ
        top = tk.Frame(w, bg=PRIMARY)
        top.pack(fill="x")
        tk.Label(top, text=session.exam.name, bg=PRIMARY, fg="white",
                 font=("Segoe UI", 11, "bold")).pack(side="left", padx=10, pady=4)
        self.project_lbl = tk.Label(top, bg=PRIMARY, fg="white", font=FONT)
        self.project_lbl.pack(side="left", padx=20)
        self.timer_lbl = tk.Label(top, bg=PRIMARY, fg="white", font=("Consolas", 14, "bold"))
        self.timer_lbl.pack(side="right", padx=12)
        tk.Label(top, text="LUYỆN TẬP" if self.training else "THI THỬ", bg="white", fg=PRIMARY,
                 font=FONT_B, padx=6).pack(side="right")

        # --- hàng nút (pack trước phần thân để không bị che khi cửa sổ thấp)
        bar = ttk.Frame(w, padding=(10, 0, 10, 8))
        bar.pack(side="bottom", fill="x")
        ttk.Button(bar, text="📂 Mở file", command=self.open_file).pack(side="left")
        ttk.Button(bar, text="↺ Làm lại dự án", command=self.reset_project).pack(side="left", padx=4)
        if self.training:
            ttk.Button(bar, text="✔ Kiểm tra dự án", command=self.check_current).pack(side="left", padx=4)
        ttk.Button(bar, text="Nộp bài", style="Accent.TButton", command=self.submit).pack(side="right")
        self.next_btn = ttk.Button(bar, text="Dự án sau ▶", command=lambda: self.go(self.idx + 1))
        self.next_btn.pack(side="right", padx=4)
        self.prev_btn = ttk.Button(bar, text="◀ Dự án trước", command=lambda: self.go(self.idx - 1))
        self.prev_btn.pack(side="right")

        # --- thân: mô tả + danh sách nhiệm vụ (cuộn được)
        body = ttk.Frame(w, padding=(10, 6))
        body.pack(fill="both", expand=True)
        self.intro_lbl = ttk.Label(body, font=("Segoe UI", 10, "italic"))
        self.intro_lbl.pack(anchor="w")

        holder = ttk.Frame(body)
        holder.pack(fill="both", expand=True, pady=4)
        self.canvas = tk.Canvas(holder, highlightthickness=0)
        sb = ttk.Scrollbar(holder, orient="vertical", command=self.canvas.yview)
        self.tasks_frame = ttk.Frame(self.canvas)
        self.tasks_frame.bind("<Configure>",
                              lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.tasks_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=sb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        w.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(-1 if e.delta > 0 else 1, "units"))

        self.show_project()
        self.open_file()
        self.tick()

    # ------------------------------------------------------------ hiển thị
    def show_project(self) -> None:
        project = self.s.exam.projects[self.idx]
        n = len(self.s.exam.projects)
        self.project_lbl.config(text=f"Dự án {self.idx + 1}/{n}:  {project.name}   ({project.filename})")
        self.intro_lbl.config(text=project.intro + "  Làm xong nhớ LƯU file (Ctrl+S).")
        self.prev_btn.state(["disabled"] if self.idx == 0 else ["!disabled"])
        self.next_btn.state(["disabled"] if self.idx == n - 1 else ["!disabled"])

        for child in self.tasks_frame.winfo_children():
            child.destroy()
        self.status_lbls = []
        self.task_vars = []
        for t_idx, task in enumerate(project.tasks):
            key = (self.idx, t_idx)
            row = ttk.Frame(self.tasks_frame)
            row.pack(fill="x", pady=1)
            done = tk.BooleanVar(value=key in self.s.done)
            mark = tk.BooleanVar(value=key in self.s.marked)
            self.task_vars.append((done, mark))
            ttk.Label(row, text=f"Nhiệm vụ {t_idx + 1}", font=FONT_B, width=11).pack(side="left")
            ttk.Checkbutton(row, text="Đã làm", variable=done,
                            command=lambda k=key, v=done: self._toggle(self.s.done, k, v)).pack(side="left")
            ttk.Checkbutton(row, text="Đánh dấu", variable=mark,
                            command=lambda k=key, v=mark: self._toggle(self.s.marked, k, v)).pack(side="left")
            if self.training:
                ttk.Button(row, text="Gợi ý", width=6,
                           command=lambda t=task: messagebox.showinfo("Gợi ý", t.hint, parent=self.win)
                           ).pack(side="left", padx=4)
            status = tk.Label(row, text="", width=3, font=FONT_B)
            status.pack(side="left")
            self.status_lbls.append(status)
            ttk.Label(row, text=task.title, font=FONT, wraplength=self.win.winfo_screenwidth() - 420
                      ).pack(side="left", padx=6)
        self.canvas.yview_moveto(0)

    @staticmethod
    def _toggle(store: set, key, var: tk.BooleanVar) -> None:
        (store.add if var.get() else store.discard)(key)

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
        m, sec = divmod(max(self.seconds, 0), 60)
        self.timer_lbl.config(text=f"{m:02d}:{sec:02d}",
                              fg="#ffcdd2" if not self.training and self.seconds < 300 else "white")
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
        if 0 <= idx < len(self.s.exam.projects):
            messagebox.showinfo("Chuyển dự án",
                                "Nhớ LƯU (Ctrl+S) file hiện tại trước khi sang dự án khác.", parent=self.win)
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
        for lbl, r in zip(self.status_lbls, results):
            lbl.config(text="✔" if r.correct else "✘", fg=OK_COLOR if r.correct else BAD_COLOR)
        ok = sum(r.correct for r in results)
        errs = [r.error for r in results if r.error]
        extra = f"\n\nLỗi đọc file: {errs[0]}" if errs else ""
        messagebox.showinfo("Kết quả dự án", f"Đúng {ok}/{len(results)} nhiệm vụ.{extra}", parent=self.win)

    def submit(self) -> None:
        opened = [i for i in range(len(self.s.exam.projects)) if core.file_is_open(self.s.file_of(i))]
        msg = "Nộp bài và chấm điểm?"
        if opened:
            msg = ("Vẫn còn file đang mở trong Office. Chỉ phần đã LƯU mới được chấm.\n\n" + msg)
        if self.s.marked:
            msg = f"Bạn còn {len(self.s.marked)} nhiệm vụ đánh dấu xem lại.\n" + msg
        if messagebox.askyesno("Nộp bài", msg, parent=self.win):
            self.finish()

    def finish(self) -> None:
        self.finished = True
        report = core.grade(self.s)
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
        w = self.win = tk.Toplevel(root)
        w.title("Kết quả")
        w.geometry("900x560")
        w.protocol("WM_DELETE_WINDOW", self.close)

        head = ttk.Frame(w, padding=16)
        head.pack(fill="x")
        color = OK_COLOR if report["passed"] else BAD_COLOR
        tk.Label(head, text=f'{report["score"]}', font=("Segoe UI", 40, "bold"), fg=color).pack(side="left")
        info = ttk.Frame(head)
        info.pack(side="left", padx=16)
        tk.Label(info, text="ĐẠT" if report["passed"] else "CHƯA ĐẠT", font=("Segoe UI", 18, "bold"),
                 fg=color).pack(anchor="w")
        ttk.Label(info, text=f'/1000 điểm  •  Điểm đạt: {core.PASS_SCORE}  •  '
                             f'Đúng {report["correct"]}/{report["total"]} nhiệm vụ',
                  font=FONT).pack(anchor="w")

        cols = ("project", "task", "result")
        tree = self.tree = ttk.Treeview(w, columns=cols, show="headings", height=12)
        for c, text, width in zip(cols, ("Dự án", "Nhiệm vụ", "Kết quả"), (200, 580, 80)):
            tree.heading(c, text=text)
            tree.column(c, width=width, anchor="w")
        tree.tag_configure("ok", foreground=OK_COLOR)
        tree.tag_configure("bad", foreground=BAD_COLOR)
        self.results = report["results"]
        for i, r in enumerate(self.results):
            tree.insert("", "end", iid=str(i), values=(r.project, r.task, "Đúng" if r.correct else "Sai"),
                        tags=("ok" if r.correct else "bad",))
        tree.pack(fill="both", expand=True, padx=16)
        tree.bind("<<TreeviewSelect>>", self.on_select)

        self.hint = tk.Text(w, height=4, wrap="word", font=FONT, relief="flat", background="#f4f6f8")
        self.hint.pack(fill="x", padx=16, pady=8)
        self._set_hint("Chọn một nhiệm vụ để xem cách làm.")

        btns = ttk.Frame(w, padding=(16, 0, 16, 12))
        btns.pack(fill="x")
        ttk.Button(btns, text="Mở thư mục bài làm",
                   command=lambda: core.open_in_office(session.workdir)).pack(side="left")
        ttk.Button(btns, text="Về màn hình chính", style="Accent.TButton", command=self.close).pack(side="right")

    def _set_hint(self, text: str) -> None:
        self.hint.config(state="normal")
        self.hint.delete("1.0", "end")
        self.hint.insert("1.0", text)
        self.hint.config(state="disabled")

    def on_select(self, _event) -> None:
        sel = self.tree.selection()
        if sel:
            r = self.results[int(sel[0])]
            text = f"Cách làm: {r.hint}"
            if r.error:
                text += f"\n(Lỗi khi đọc file: {r.error})"
            self._set_hint(text)

    def close(self) -> None:
        self.win.destroy()
        self.root.deiconify()


def main() -> None:
    root = tk.Tk()
    LauncherWindow(root)
    root.mainloop()
