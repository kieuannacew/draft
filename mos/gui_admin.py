"""Màn hình Quản trị (chỉ tài khoản quản trị mới mở được).

  * Tab "Tài khoản"  – tạo / sửa / khóa / xóa tài khoản, đặt lại mật khẩu,
                       tạo hàng loạt tài khoản cho cả lớp.
  * Tab "Đề thi"     – danh sách đề tự soạn; soạn đề mới / sửa / kiểm tra.
  * Tab "Kết quả"    – kết quả làm bài của mọi học viên, xuất CSV.
  * ExamEditor       – form soạn đề: không cần viết de.json bằng tay.
"""
from __future__ import annotations

import csv
import secrets
import shutil
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from . import core, custom, rules
from .accounts import ROLE_NAMES, STUDENT, AccountError, AccountStore
from .ui import (BG, BORDER, CARD, DANGER, MUTED, PRIMARY, SUCCESS, TEXT, WARN, WARN_BG, Btn, F, card,
                 entry, fmt_time, label, make_tree, modal, pill, text_box)

MON_NAMES = {"WORD": "Word", "EXCEL": "Excel", "POWERPOINT": "PowerPoint"}
FILE_TYPES = {"WORD": [("Word", "*.docx")], "EXCEL": [("Excel", "*.xlsx")],
              "POWERPOINT": [("PowerPoint", "*.pptx")]}


def _toolbar(parent):
    bar = tk.Frame(parent, bg=BG)
    bar.pack(fill="x", pady=(0, 10))
    return bar


def _form_row(parent, text, widget_factory, row, hint=None):
    label(parent, text, bold=True, size=9).grid(row=row, column=0, sticky="w", pady=(8, 2))
    w = widget_factory(parent)
    w.grid(row=row + 1, column=0, sticky="ew")
    if hint:
        label(parent, hint, muted=True, size=8).grid(row=row + 2, column=0, sticky="w")
    return w


# ====================================================================== cửa sổ chính


class AdminWindow:
    def __init__(self, root, store: AccountStore, me, on_close=None):
        self.store, self.me, self.on_close = store, me, on_close
        w = self.win = tk.Toplevel(root, bg=BG)
        w.title("Quản trị – Luyện thi MOS")
        w.geometry("1040x640")
        w.minsize(900, 560)
        w.protocol("WM_DELETE_WINDOW", self.close)

        head = tk.Frame(w, bg=BG, padx=20, pady=14)
        head.pack(fill="x")
        label(head, "Quản trị", bold=True, size=18).pack(side="left")
        label(head, f"   Đăng nhập: {me.ho_ten} ({me.username})", muted=True).pack(side="left")

        st = ttk.Style(w)
        st.configure("Admin.TNotebook", background=BG, borderwidth=0)
        st.configure("Admin.TNotebook.Tab", font=F(10, "bold"), padding=(18, 8), background=BORDER,
                     foreground=MUTED)
        st.map("Admin.TNotebook.Tab", background=[("selected", CARD)], foreground=[("selected", PRIMARY)])
        nb = ttk.Notebook(w, style="Admin.TNotebook")
        nb.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        for title, builder in (("  Tài khoản  ", self._tab_accounts), ("  Đề thi  ", self._tab_exams),
                               ("  Kết quả  ", self._tab_results)):
            frame = tk.Frame(nb, bg=BG, padx=12, pady=12)
            nb.add(frame, text=title)
            builder(frame)

    def close(self):
        self.win.destroy()
        if self.on_close:
            self.on_close()

    # ------------------------------------------------------------ Tài khoản
    def _tab_accounts(self, tab):
        bar = _toolbar(tab)
        Btn(bar, "+ Thêm tài khoản", self.add_user, kind="primary").pack(side="left")
        Btn(bar, "Tạo cho cả lớp", self.bulk_users).pack(side="left", padx=6)
        Btn(bar, "Sửa", self.edit_user).pack(side="left")
        Btn(bar, "Đặt lại mật khẩu", self.reset_password).pack(side="left", padx=6)
        Btn(bar, "Khóa / Mở khóa", self.toggle_lock).pack(side="left")
        Btn(bar, "Xóa", self.delete_user).pack(side="right")
        box, self.users_tree = make_tree(tab, [
            ("user", "Tên đăng nhập", 140, False), ("name", "Họ tên", 200, True),
            ("role", "Vai trò", 90, False), ("status", "Trạng thái", 110, False),
            ("exp", "Hạn dùng", 100, False), ("count", "Số lần thi", 80, False),
            ("best", "Điểm cao nhất", 100, False)], height=14)
        box.pack(fill="both", expand=True)
        self.users_tree.bind("<Double-1>", lambda e: self.edit_user())
        self.refresh_users()

    def refresh_users(self):
        tree = self.users_tree
        tree.delete(*tree.get_children())
        history = core.load_history()
        for a in self.store.list():
            mine = [h for h in history if h.get("user") == a.username]
            status = "Bị khóa" if a.khoa else "Hết hạn" if a.expired() else "Hoạt động"
            tree.insert("", "end", iid=a.username, tags=("bad",) if status != "Hoạt động" else (),
                        values=(a.username, a.ho_ten, ROLE_NAMES[a.vai_tro], status, a.han_dung or "—",
                                len(mine), max((h["score"] for h in mine), default="—")))

    def _selected_user(self):
        sel = self.users_tree.selection()
        if not sel:
            messagebox.showinfo("Chọn tài khoản", "Hãy chọn một tài khoản trong danh sách.", parent=self.win)
            return None
        return self.store.get(sel[0])

    def add_user(self):
        UserDialog(self.win, self.store, None, self.refresh_users)

    def edit_user(self):
        acc = self._selected_user()
        if acc:
            UserDialog(self.win, self.store, acc, self.refresh_users)

    def bulk_users(self):
        BulkUsersDialog(self.win, self.store, self.refresh_users)

    def reset_password(self):
        acc = self._selected_user()
        if not acc:
            return
        pw = _random_password()
        if messagebox.askyesno("Đặt lại mật khẩu",
                               f"Đặt mật khẩu mới cho '{acc.username}' là:\n\n    {pw}\n\n"
                               "Người dùng sẽ phải đổi mật khẩu khi đăng nhập. Tiếp tục?", parent=self.win):
            self.store.set_password(acc.username, pw, must_change=True)
            self.win.clipboard_clear()
            self.win.clipboard_append(pw)
            messagebox.showinfo("Xong", f"Mật khẩu mới: {pw}\n(đã chép vào clipboard)", parent=self.win)

    def toggle_lock(self):
        acc = self._selected_user()
        if not acc:
            return
        if acc.username == self.me.username:
            messagebox.showwarning("Không thể", "Không thể khóa tài khoản đang đăng nhập.", parent=self.win)
            return
        try:
            self.store.update(acc.username, khoa=not acc.khoa)
        except AccountError as exc:
            messagebox.showerror("Lỗi", str(exc), parent=self.win)
        self.refresh_users()

    def delete_user(self):
        acc = self._selected_user()
        if not acc:
            return
        if acc.username == self.me.username:
            messagebox.showwarning("Không thể", "Không thể xóa tài khoản đang đăng nhập.", parent=self.win)
            return
        if messagebox.askyesno("Xóa tài khoản", f"Xóa tài khoản '{acc.username}'?\n"
                               "(Kết quả làm bài cũ vẫn được giữ lại.)", parent=self.win):
            try:
                self.store.delete(acc.username)
            except AccountError as exc:
                messagebox.showerror("Lỗi", str(exc), parent=self.win)
            self.refresh_users()

    # ------------------------------------------------------------ Đề thi
    def _tab_exams(self, tab):
        bar = _toolbar(tab)
        Btn(bar, "+ Soạn đề mới", self.new_exam, kind="primary").pack(side="left")
        Btn(bar, "Sửa", self.edit_exam).pack(side="left", padx=6)
        Btn(bar, "Kiểm tra đề", self.check_exam).pack(side="left")
        Btn(bar, "Mở thư mục", self.open_exam_folder).pack(side="left", padx=6)
        Btn(bar, "Xóa", self.delete_exam).pack(side="right")
        label(tab, "Đề tự soạn được gộp vào bài thi của môn tương ứng (Word / Excel / PowerPoint) "
                   "trên màn hình chính.", muted=True, size=9).pack(anchor="w", pady=(0, 8))
        box, self.exams_tree = make_tree(tab, [
            ("name", "Tên đề", 260, True), ("mon", "Môn", 100, False), ("proj", "Dự án", 70, False),
            ("tasks", "Nhiệm vụ", 80, False), ("status", "Trạng thái", 240, True)], height=14)
        box.pack(fill="both", expand=True)
        self.exams_tree.bind("<Double-1>", lambda e: self.edit_exam())
        self.exam_folders: dict[str, Path] = {}
        self.refresh_exams()

    def refresh_exams(self):
        tree = self.exams_tree
        tree.delete(*tree.get_children())
        self.exam_folders = {}
        for i, folder in enumerate(custom.list_exam_folders()):
            iid = str(i)
            self.exam_folders[iid] = folder
            try:
                data = custom.read_exam_json(folder)
            except (OSError, ValueError):
                data = {}
            exam, errors = custom.load_exam(folder)
            projects = data.get("du_an", [])
            tree.insert("", "end", iid=iid, tags=("bad",) if errors else ("ok",), values=(
                data.get("ten", folder.name), MON_NAMES.get(str(data.get("mon", "")).upper(), "?"),
                len(projects), sum(len(p.get("nhiem_vu", [])) for p in projects),
                f"Lỗi: {errors[0]}" if errors else "Sẵn sàng"))

    def _selected_exam(self):
        sel = self.exams_tree.selection()
        if not sel:
            messagebox.showinfo("Chọn đề", "Hãy chọn một đề trong danh sách.", parent=self.win)
            return None
        return self.exam_folders[sel[0]]

    def new_exam(self):
        ExamEditor(self.win, None, self.refresh_exams)

    def edit_exam(self):
        folder = self._selected_exam()
        if folder:
            ExamEditor(self.win, folder, self.refresh_exams)

    def check_exam(self):
        folder = self._selected_exam()
        if folder:
            CheckReport(self.win, folder)

    def open_exam_folder(self):
        folder = self._selected_exam() if self.exams_tree.selection() else custom.editable_dir()
        if folder:
            folder.mkdir(parents=True, exist_ok=True)
            core.open_in_office(folder)

    def delete_exam(self):
        folder = self._selected_exam()
        if folder and messagebox.askyesno("Xóa đề", f"Xóa toàn bộ thư mục đề:\n{folder}\n\n"
                                          "Không thể hoàn tác.", parent=self.win):
            shutil.rmtree(folder, ignore_errors=True)
            self.refresh_exams()

    # ------------------------------------------------------------ Kết quả
    def _tab_results(self, tab):
        bar = _toolbar(tab)
        label(bar, "Học viên:", bold=True).pack(side="left")
        self.filter_var = tk.StringVar(value="(Tất cả)")
        users = ["(Tất cả)"] + [a.username for a in self.store.list()]
        cb = ttk.Combobox(bar, textvariable=self.filter_var, values=users, state="readonly", width=24)
        cb.pack(side="left", padx=8)
        cb.bind("<<ComboboxSelected>>", lambda e: self.refresh_results())
        Btn(bar, "Làm mới", self.refresh_results).pack(side="left")
        Btn(bar, "Xuất CSV (Excel)", self.export_results).pack(side="right")
        self.summary = label(tab, "", muted=True)
        self.summary.pack(anchor="w", pady=(0, 8))
        box, self.results_tree = make_tree(tab, [
            ("time", "Thời gian", 145, False), ("user", "Tài khoản", 110, False),
            ("name", "Họ tên", 170, True), ("exam", "Bài thi", 200, True), ("mode", "Chế độ", 80, False),
            ("dur", "Làm trong", 80, False), ("score", "Điểm", 80, False), ("result", "Kết quả", 80, False)],
            height=14)
        box.pack(fill="both", expand=True)
        self.refresh_results()

    def _result_rows(self):
        who = self.filter_var.get()
        history = core.load_history(None if who == "(Tất cả)" else who)
        rows = []
        for h in reversed(history):
            acc = self.store.get(h["user"]) if h.get("user") else None
            rows.append((h["time"], h.get("user") or "—", acc.ho_ten if acc else "—", h["exam"],
                         "Luyện tập" if h["mode"] == "training" else "Thi thử",
                         fmt_time(h["duration"]) if h.get("duration") is not None else "—", h["score"], "Đạt" if h["passed"] else "Chưa đạt"))
        return rows

    def refresh_results(self):
        tree = self.results_tree
        tree.delete(*tree.get_children())
        rows = self._result_rows()
        for r in rows:
            tree.insert("", "end", values=r, tags=("ok",) if r[-1] == "Đạt" else ("bad",))
        passed = sum(r[-1] == "Đạt" for r in rows)
        self.summary.config(text=f"{len(rows)} lượt làm bài  •  {passed} lượt đạt  •  điểm trung bình "
                                 f"{round(sum(r[6] for r in rows) / len(rows)) if rows else '—'}")

    def export_results(self):
        path = filedialog.asksaveasfilename(parent=self.win, defaultextension=".csv",
                                            initialfile="ket_qua_MOS.csv", filetypes=[("CSV", "*.csv")])
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:   # utf-8-sig để Excel đọc đúng tiếng Việt
            w = csv.writer(f)
            w.writerow(["Thời gian", "Tài khoản", "Họ tên", "Bài thi", "Chế độ", "Làm trong", "Điểm", "Kết quả"])
            w.writerows(self._result_rows())
        messagebox.showinfo("Đã xuất", f"Đã lưu:\n{path}", parent=self.win)


def _random_password() -> str:
    alphabet = "abcdefghjkmnpqrstuvwxyz23456789"
    return "".join(secrets.choice(alphabet) for _ in range(6))


# ====================================================================== hộp thoại tài khoản


class UserDialog:
    def __init__(self, parent, store: AccountStore, acc, on_done):
        self.store, self.acc, self.on_done = store, acc, on_done
        w = self.win = tk.Toplevel(parent, bg=BG, padx=24, pady=18)
        w.title("Sửa tài khoản" if acc else "Thêm tài khoản")
        w.resizable(False, False)
        form = tk.Frame(w, bg=BG)
        form.pack(fill="x")
        form.columnconfigure(0, weight=1, minsize=340)

        self.username = tk.StringVar(value=acc.username if acc else "")
        self.ho_ten = tk.StringVar(value=acc.ho_ten if acc else "")
        self.password = tk.StringVar(value="" if acc else _random_password())
        self.role = tk.StringVar(value=ROLE_NAMES[acc.vai_tro if acc else STUDENT])
        self.han = tk.StringVar(value=(acc.han_dung or "") if acc else "")
        self.must = tk.BooleanVar(value=True)

        e = _form_row(form, "Tên đăng nhập", lambda p: entry(p, self.username), 0,
                      "chữ thường không dấu, số, _ hoặc .  (vd: nguyenvana)")
        if acc:
            e.configure(state="disabled")
        _form_row(form, "Họ tên", lambda p: entry(p, self.ho_ten), 3)
        if not acc:
            _form_row(form, "Mật khẩu", lambda p: entry(p, self.password), 5, "đã tạo ngẫu nhiên, có thể sửa")
        _form_row(form, "Vai trò", lambda p: ttk.Combobox(p, textvariable=self.role, state="readonly",
                                                          values=list(ROLE_NAMES.values())), 8)
        _form_row(form, "Hạn dùng (tuỳ chọn)", lambda p: entry(p, self.han), 10,
                  "YYYY-MM-DD, vd 2026-12-31. Để trống = không giới hạn")
        if not acc:
            tk.Checkbutton(form, text="Bắt đổi mật khẩu ở lần đăng nhập đầu", variable=self.must, bg=BG,
                           font=F(9), activebackground=BG).grid(row=13, column=0, sticky="w", pady=(8, 0))

        bar = tk.Frame(w, bg=BG)
        bar.pack(fill="x", pady=(18, 0))
        Btn(bar, "Lưu", self.save, kind="primary").pack(side="right")
        Btn(bar, "Hủy", w.destroy).pack(side="right", padx=8)
        modal(w, parent)

    def save(self):
        role = next(k for k, v in ROLE_NAMES.items() if v == self.role.get())
        try:
            if self.acc:
                self.store.update(self.acc.username, ho_ten=self.ho_ten.get(), vai_tro=role,
                                  han_dung=self.han.get())
            else:
                acc = self.store.create(self.username.get(), self.ho_ten.get(), self.password.get(), role,
                                        han_dung=self.han.get(), must_change=self.must.get())
                messagebox.showinfo("Đã tạo tài khoản",
                                    f"Tên đăng nhập:  {acc.username}\nMật khẩu:  {self.password.get()}\n\n"
                                    "Hãy gửi thông tin này cho người dùng.", parent=self.win)
        except AccountError as exc:
            messagebox.showerror("Lỗi", str(exc), parent=self.win)
            return
        self.win.destroy()
        self.on_done()


class BulkUsersDialog:
    """Tạo nhiều tài khoản một lúc từ danh sách "tên đăng nhập, họ tên"."""

    def __init__(self, parent, store: AccountStore, on_done):
        self.store, self.on_done, self.parent = store, on_done, parent
        w = self.win = tk.Toplevel(parent, bg=BG, padx=24, pady=18)
        w.title("Tạo tài khoản cho cả lớp")
        label(w, "Mỗi dòng một học viên:  tên_đăng_nhập, Họ tên", bold=True).pack(anchor="w")
        label(w, "Ví dụ:   nguyenvana, Nguyễn Văn A\nMật khẩu được tạo ngẫu nhiên cho từng người.",
              muted=True, size=9).pack(anchor="w", pady=(2, 8))
        self.text = text_box(w, height=12)
        self.text.configure(width=56)
        self.text.pack(fill="both", expand=True)
        row = tk.Frame(w, bg=BG)
        row.pack(fill="x", pady=(10, 0))
        label(row, "Hạn dùng (tuỳ chọn):").pack(side="left")
        self.han = tk.StringVar()
        entry(row, self.han, width=14).pack(side="left", padx=8)
        label(row, "YYYY-MM-DD", muted=True, size=9).pack(side="left")
        bar = tk.Frame(w, bg=BG)
        bar.pack(fill="x", pady=(16, 0))
        Btn(bar, "Tạo tài khoản", self.create, kind="primary").pack(side="right")
        Btn(bar, "Hủy", w.destroy).pack(side="right", padx=8)
        modal(w, parent)

    def create(self):
        created, errors = [], []
        for n, line in enumerate(self.text.get("1.0", "end").splitlines(), start=1):
            if not line.strip():
                continue
            username, _, name = line.partition(",")
            pw = _random_password()
            try:
                acc = self.store.create(username, name, pw, STUDENT, han_dung=self.han.get())
                created.append((acc.username, acc.ho_ten, pw))
            except AccountError as exc:
                errors.append(f"Dòng {n}: {exc}")
        self.win.destroy()
        self.on_done()
        if created:
            CredentialsDialog(self.parent, created, errors)
        elif errors:
            messagebox.showerror("Không tạo được", "\n".join(errors[:15]), parent=self.parent)


class CredentialsDialog:
    """Hiện danh sách tài khoản vừa tạo và cho lưu ra file để phát cho học viên."""

    def __init__(self, parent, rows, errors):
        self.rows = rows
        w = self.win = tk.Toplevel(parent, bg=BG, padx=24, pady=18)
        w.title("Tài khoản vừa tạo")
        label(w, f"Đã tạo {len(rows)} tài khoản. Hãy lưu danh sách để phát cho học viên:", bold=True
              ).pack(anchor="w", pady=(0, 8))
        box, tree = make_tree(w, [("u", "Tên đăng nhập", 150, False), ("n", "Họ tên", 220, True),
                                  ("p", "Mật khẩu", 110, False)], height=min(12, len(rows)))
        box.pack(fill="both", expand=True)
        for r in rows:
            tree.insert("", "end", values=r)
        if errors:
            label(w, "Bỏ qua:\n" + "\n".join(errors[:8]), size=9, fg=DANGER).pack(anchor="w", pady=(8, 0))
        bar = tk.Frame(w, bg=BG)
        bar.pack(fill="x", pady=(16, 0))
        Btn(bar, "Lưu danh sách (CSV)", self.save, kind="primary").pack(side="right")
        Btn(bar, "Đóng", w.destroy).pack(side="right", padx=8)
        modal(w, parent)

    def save(self):
        path = filedialog.asksaveasfilename(parent=self.win, defaultextension=".csv",
                                            initialfile="tai_khoan_hoc_vien.csv", filetypes=[("CSV", "*.csv")])
        if path:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(["Tên đăng nhập", "Họ tên", "Mật khẩu"])
                w.writerows(self.rows)
            messagebox.showinfo("Đã lưu", path, parent=self.win)


class ChangePasswordDialog:
    def __init__(self, parent, store: AccountStore, acc, forced=False, on_done=None):
        self.store, self.acc, self.forced, self.on_done = store, acc, forced, on_done
        self.ok = False
        w = self.win = tk.Toplevel(parent, bg=BG, padx=24, pady=18)
        w.title("Đổi mật khẩu")
        w.resizable(False, False)
        if forced:
            tk.Label(w, text="Đây là lần đăng nhập đầu tiên — hãy đặt mật khẩu mới.", bg=WARN_BG, fg=WARN,
                     font=F(9, "bold"), padx=10, pady=6).pack(fill="x", pady=(0, 6))
        form = tk.Frame(w, bg=BG)
        form.pack(fill="x")
        form.columnconfigure(0, weight=1, minsize=320)
        self.old, self.new, self.again = tk.StringVar(), tk.StringVar(), tk.StringVar()
        r = 0
        if not forced:
            _form_row(form, "Mật khẩu hiện tại", lambda p: entry(p, self.old, show="•"), r)
            r = 2
        first = _form_row(form, "Mật khẩu mới", lambda p: entry(p, self.new, show="•"), r)
        _form_row(form, "Nhập lại mật khẩu mới", lambda p: entry(p, self.again, show="•"), r + 2)
        bar = tk.Frame(w, bg=BG)
        bar.pack(fill="x", pady=(18, 0))
        Btn(bar, "Lưu", self.save, kind="primary").pack(side="right")
        if not forced:
            Btn(bar, "Hủy", w.destroy).pack(side="right", padx=8)
        else:
            w.protocol("WM_DELETE_WINDOW", lambda: None)
        w.bind("<Return>", lambda e: self.save())
        modal(w, parent)
        first.focus_set()

    def save(self):
        try:
            if not self.forced:
                self.store.authenticate(self.acc.username, self.old.get())
            if self.new.get() != self.again.get():
                raise AccountError("Hai lần nhập mật khẩu mới không giống nhau.")
            if self.forced and self.new.get() == "admin":
                raise AccountError("Hãy chọn mật khẩu khác mật khẩu mặc định.")
            self.store.set_password(self.acc.username, self.new.get())
        except AccountError as exc:
            messagebox.showerror("Lỗi", str(exc), parent=self.win)
            return
        self.ok = True
        self.win.destroy()
        if self.on_done:
            self.on_done()


# ====================================================================== Soạn đề


class ExamEditor:
    """Form soạn đề: Đề → Dự án (file gốc) → Nhiệm vụ → Luật chấm."""

    def __init__(self, parent, folder: Path | None, on_saved):
        self.parent, self.folder, self.on_saved = parent, folder, on_saved
        if folder:
            self.data = custom.read_exam_json(folder)
        else:
            self.data = {"mon": "EXCEL", "ten": "", "du_an": []}
        self.p_idx = self.t_idx = None
        self.dirty = False

        w = self.win = tk.Toplevel(parent, bg=BG)
        w.title("Soạn đề" + (f" – {self.data.get('ten', '')}" if folder else " mới"))
        w.geometry("1180x720")
        w.minsize(1000, 620)
        w.protocol("WM_DELETE_WINDOW", self.close)

        # --- thông tin đề
        top = tk.Frame(w, bg=BG, padx=16, pady=12)
        top.pack(fill="x")
        label(top, "Tên đề", bold=True).pack(side="left")
        self.ten = tk.StringVar(value=self.data.get("ten", ""))
        entry(top, self.ten, width=40).pack(side="left", padx=8, ipady=3)
        label(top, "Môn", bold=True).pack(side="left", padx=(16, 0))
        self.mon = tk.StringVar(value=MON_NAMES.get(str(self.data.get("mon", "EXCEL")).upper(), "Excel"))
        cb = ttk.Combobox(top, textvariable=self.mon, values=list(MON_NAMES.values()), state="readonly", width=12)
        cb.pack(side="left", padx=8)
        cb.bind("<<ComboboxSelected>>", lambda e: self._mark())
        Btn(top, "Lưu đề", self.save, kind="primary").pack(side="right")
        Btn(top, "Lưu & kiểm tra", self.save_and_check).pack(side="right", padx=8)

        # --- bottom help
        tk.Label(w, text="Cách soạn: ① Thêm dự án và chọn file gốc (file chưa làm)  ② Thêm nhiệm vụ, ghi yêu cầu "
                         "và gợi ý  ③ Thêm luật chấm  ④ Chọn file đáp án (đã làm đúng) rồi bấm “Lưu & kiểm tra”.",
                 bg=WARN_BG, fg="#78350f", font=F(9), anchor="w", padx=12, pady=6).pack(side="bottom", fill="x")

        body = tk.Frame(w, bg=BG, padx=16)
        body.pack(fill="both", expand=True, pady=(0, 10))
        body.columnconfigure(0, weight=2, uniform="c")
        body.columnconfigure(1, weight=3, uniform="c")
        body.columnconfigure(2, weight=4, uniform="c")
        body.rowconfigure(0, weight=1)

        # cột 1: dự án
        c1 = card(body, padx=12, pady=10)
        c1.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        label(c1, "① Dự án", bold=True, size=11).pack(anchor="w")
        self.p_list = self._listbox(c1, self.on_project_select)
        bar = tk.Frame(c1, bg=CARD)
        bar.pack(fill="x", pady=(8, 0))
        Btn(bar, "+ Thêm", self.add_project, kind="ghost", padx=10).pack(side="left")
        Btn(bar, "↑", lambda: self.move_project(-1), padx=8).pack(side="left", padx=4)
        Btn(bar, "↓", lambda: self.move_project(1), padx=8).pack(side="left")
        Btn(bar, "Xóa", self.delete_project, padx=10).pack(side="right")

        # cột 2: chi tiết dự án + danh sách nhiệm vụ
        c2 = self.c2 = card(body, padx=12, pady=10)
        c2.grid(row=0, column=1, sticky="nsew", padx=(0, 8))
        label(c2, "Thông tin dự án", bold=True, size=11).pack(anchor="w")
        self.p_ten, self.p_mota = tk.StringVar(), tk.StringVar()
        label(c2, "Tên dự án", size=9, muted=True).pack(anchor="w", pady=(6, 0))
        entry(c2, self.p_ten).pack(fill="x", ipady=2)
        label(c2, "Mô tả tình huống", size=9, muted=True).pack(anchor="w", pady=(6, 0))
        entry(c2, self.p_mota).pack(fill="x", ipady=2)
        for var in (self.p_ten, self.p_mota):
            var.trace_add("write", lambda *a: self._commit_project())
        files = tk.Frame(c2, bg=CARD)
        files.pack(fill="x", pady=(8, 0))
        self.src_lbl = label(files, "", size=9)
        self.src_lbl.grid(row=0, column=0, sticky="w")
        Btn(files, "Chọn file gốc…", self.pick_source, padx=8, pady=3, size=9).grid(row=0, column=1, sticky="e")
        self.ans_lbl = label(files, "", size=9)
        self.ans_lbl.grid(row=1, column=0, sticky="w", pady=(4, 0))
        Btn(files, "Chọn file đáp án…", self.pick_answer, padx=8, pady=3, size=9).grid(
            row=1, column=1, sticky="e", pady=(4, 0))
        files.columnconfigure(0, weight=1)

        label(c2, "② Nhiệm vụ", bold=True, size=11).pack(anchor="w", pady=(12, 0))
        self.t_list = self._listbox(c2, self.on_task_select)
        bar = tk.Frame(c2, bg=CARD)
        bar.pack(fill="x", pady=(8, 0))
        Btn(bar, "+ Thêm", self.add_task, kind="ghost", padx=10).pack(side="left")
        Btn(bar, "↑", lambda: self.move_task(-1), padx=8).pack(side="left", padx=4)
        Btn(bar, "↓", lambda: self.move_task(1), padx=8).pack(side="left")
        Btn(bar, "Xóa", self.delete_task, padx=10).pack(side="right")

        # cột 3: chi tiết nhiệm vụ + luật chấm
        c3 = self.c3 = card(body, padx=12, pady=10)
        c3.grid(row=0, column=2, sticky="nsew")
        label(c3, "Chi tiết nhiệm vụ", bold=True, size=11).pack(anchor="w")
        label(c3, "Yêu cầu (người học nhìn thấy)", size=9, muted=True).pack(anchor="w", pady=(6, 0))
        self.t_yc = text_box(c3, height=3)
        self.t_yc.pack(fill="x")
        label(c3, "Gợi ý cách làm", size=9, muted=True).pack(anchor="w", pady=(6, 0))
        self.t_gy = text_box(c3, height=3)
        self.t_gy.pack(fill="x")
        for t in (self.t_yc, self.t_gy):
            t.bind("<KeyRelease>", lambda e: self._commit_task())
        label(c3, "③ Luật chấm (phải đúng tất cả)", bold=True, size=11).pack(anchor="w", pady=(12, 0))
        self.r_list = self._listbox(c3, None, height=6)
        self.r_list.bind("<Double-1>", lambda e: self.edit_rule())
        bar = tk.Frame(c3, bg=CARD)
        bar.pack(fill="x", pady=(8, 0))
        Btn(bar, "+ Thêm luật", self.add_rule, kind="ghost", padx=10).pack(side="left")
        Btn(bar, "Sửa", self.edit_rule, padx=10).pack(side="left", padx=4)
        Btn(bar, "Xóa", self.delete_rule, padx=10).pack(side="right")

        self.refresh_projects()
        self._show_project(0 if self.data["du_an"] else None)
        modal(w, parent)

    # ------------------------------------------------------------ tiện ích
    @staticmethod
    def _listbox(parent, on_select, height=8):
        lb = tk.Listbox(parent, height=height, font=F(10), relief="flat", bg=CARD, fg=TEXT,
                        selectbackground="#dbeafe", selectforeground=TEXT, activestyle="none",
                        highlightthickness=1, highlightbackground=BORDER, exportselection=False)
        lb.pack(fill="both", expand=True, pady=(6, 0))
        if on_select:
            lb.bind("<<ListboxSelect>>", lambda e: on_select())
        return lb

    def _mark(self):
        self.dirty = True

    def _mon_code(self):
        return next(k for k, v in MON_NAMES.items() if v == self.mon.get())

    @property
    def project(self):
        return self.data["du_an"][self.p_idx] if self.p_idx is not None else None

    @property
    def task(self):
        p = self.project
        return p["nhiem_vu"][self.t_idx] if p is not None and self.t_idx is not None else None

    @staticmethod
    def _set_state(frame, enabled):
        for child in frame.winfo_children():
            try:
                child.configure(state="normal" if enabled else "disabled")
            except tk.TclError:
                pass
            ExamEditor._set_state(child, enabled)

    # ------------------------------------------------------------ dự án
    def refresh_projects(self):
        self.p_list.delete(0, "end")
        for i, p in enumerate(self.data["du_an"], start=1):
            self.p_list.insert("end", f"{i}. {p.get('ten') or '(chưa đặt tên)'}")
        if self.p_idx is not None:
            self.p_list.selection_set(self.p_idx)

    def _show_project(self, idx):
        self.p_idx = idx
        self._loading = True
        p = self.project
        self.p_ten.set(p.get("ten", "") if p else "")
        self.p_mota.set(p.get("mo_ta", "") if p else "")
        self._loading = False
        self._refresh_files()
        self._set_state(self.c2, p is not None)
        self.refresh_tasks()
        self._show_task(0 if p and p["nhiem_vu"] else None)
        self.refresh_projects()

    def _refresh_files(self):
        p = self.project
        if not p:
            self.src_lbl.config(text="")
            self.ans_lbl.config(text="")
            return
        src = p.get("_src") or (self.folder / p["file"] if self.folder and p.get("file") else None)
        self.src_lbl.config(text=f"File gốc: {Path(src).name if src else 'chưa chọn'}",
                            fg=TEXT if src else DANGER)
        ans = p.get("_ans") or (self.folder / "dap_an" / p["file"] if self.folder and p.get("file") else None)
        has_ans = bool(ans) and Path(ans).is_file()
        self.ans_lbl.config(text=f"Đáp án: {Path(ans).name if has_ans else 'chưa có (nên có để kiểm tra)'}",
                            fg=SUCCESS if has_ans else MUTED)

    def _commit_project(self):
        if getattr(self, "_loading", False) or not self.project:
            return
        self.project["ten"] = self.p_ten.get().strip()
        self.project["mo_ta"] = self.p_mota.get().strip()
        self._mark()
        self.p_list.delete(self.p_idx)
        self.p_list.insert(self.p_idx, f"{self.p_idx + 1}. {self.project['ten'] or '(chưa đặt tên)'}")
        self.p_list.selection_set(self.p_idx)

    def on_project_select(self):
        sel = self.p_list.curselection()
        if sel and sel[0] != self.p_idx:
            self._show_project(sel[0])

    def add_project(self):
        n = len(self.data["du_an"]) + 1
        self.data["du_an"].append({"ten": f"Dự án {n}", "file": "", "mo_ta": "", "nhiem_vu": []})
        self._mark()
        self._show_project(n - 1)
        self.pick_source()

    def delete_project(self):
        if self.project and messagebox.askyesno("Xóa dự án", f"Xóa dự án “{self.project['ten']}”?",
                                                parent=self.win):
            del self.data["du_an"][self.p_idx]
            self._mark()
            self._show_project(min(self.p_idx, len(self.data["du_an"]) - 1) if self.data["du_an"] else None)

    def move_project(self, step):
        lst, i = self.data["du_an"], self.p_idx
        if i is not None and 0 <= i + step < len(lst):
            lst[i], lst[i + step] = lst[i + step], lst[i]
            self._mark()
            self._show_project(i + step)

    def pick_source(self):
        if not self.project:
            return
        path = filedialog.askopenfilename(parent=self.win, title="Chọn file gốc (file chưa làm)",
                                          filetypes=FILE_TYPES[self._mon_code()])
        if path:
            self.project["_src"] = path
            self.project["file"] = Path(path).name
            self._mark()
            self._refresh_files()

    def pick_answer(self):
        if not self.project:
            return
        path = filedialog.askopenfilename(parent=self.win, title="Chọn file đáp án (đã làm đúng hết)",
                                          filetypes=FILE_TYPES[self._mon_code()])
        if path:
            self.project["_ans"] = path
            self._mark()
            self._refresh_files()

    # ------------------------------------------------------------ nhiệm vụ
    def refresh_tasks(self):
        self.t_list.delete(0, "end")
        for i, t in enumerate(self.project["nhiem_vu"] if self.project else [], start=1):
            self.t_list.insert("end", f"{i}. {t.get('yeu_cau') or '(chưa có yêu cầu)'}")
        if self.t_idx is not None:
            self.t_list.selection_set(self.t_idx)

    def _show_task(self, idx):
        self.t_idx = idx
        t = self.task
        for box, key in ((self.t_yc, "yeu_cau"), (self.t_gy, "goi_y")):
            box.configure(state="normal")
            box.delete("1.0", "end")
            box.insert("1.0", t.get(key, "") if t else "")
        self._set_state(self.c3, t is not None)
        self.refresh_rules()
        self.refresh_tasks()

    def _commit_task(self):
        t = self.task
        if not t:
            return
        t["yeu_cau"] = self.t_yc.get("1.0", "end").strip()
        t["goi_y"] = self.t_gy.get("1.0", "end").strip()
        self._mark()
        self.t_list.delete(self.t_idx)
        self.t_list.insert(self.t_idx, f"{self.t_idx + 1}. {t['yeu_cau'] or '(chưa có yêu cầu)'}")
        self.t_list.selection_set(self.t_idx)

    def on_task_select(self):
        sel = self.t_list.curselection()
        if sel and sel[0] != self.t_idx:
            self._show_task(sel[0])

    def add_task(self):
        if not self.project:
            return
        self.project["nhiem_vu"].append({"yeu_cau": "", "goi_y": "", "cham": []})
        self._mark()
        self._show_task(len(self.project["nhiem_vu"]) - 1)
        self.t_yc.focus_set()

    def delete_task(self):
        if self.task and messagebox.askyesno("Xóa nhiệm vụ", "Xóa nhiệm vụ đang chọn?", parent=self.win):
            tasks = self.project["nhiem_vu"]
            del tasks[self.t_idx]
            self._mark()
            self._show_task(min(self.t_idx, len(tasks) - 1) if tasks else None)

    def move_task(self, step):
        if not self.project or self.t_idx is None:
            return
        lst, i = self.project["nhiem_vu"], self.t_idx
        if 0 <= i + step < len(lst):
            lst[i], lst[i + step] = lst[i + step], lst[i]
            self._mark()
            self._show_task(i + step)

    # ------------------------------------------------------------ luật chấm
    def _rules(self):
        t = self.task
        if t is None:
            return []
        t["cham"] = rules._as_list(t.get("cham") or [])
        return t["cham"]

    def refresh_rules(self):
        self.r_list.delete(0, "end")
        for spec in self._rules():
            self.r_list.insert("end", describe_spec(spec))

    def add_rule(self):
        if self.task is not None:
            RuleDialog(self.win, self._mon_code(), None, self._rule_saved(None))

    def edit_rule(self):
        sel = self.r_list.curselection()
        if sel:
            RuleDialog(self.win, self._mon_code(), self._rules()[sel[0]], self._rule_saved(sel[0]))

    def _rule_saved(self, idx):
        def done(spec):
            specs = self._rules()
            if idx is None:
                specs.append(spec)
            else:
                specs[idx] = spec
            self._mark()
            self.refresh_rules()
        return done

    def delete_rule(self):
        sel = self.r_list.curselection()
        if sel:
            del self._rules()[sel[0]]
            self._mark()
            self.refresh_rules()

    # ------------------------------------------------------------ lưu
    def _validate(self) -> list[str]:
        problems = []
        if not self.ten.get().strip():
            problems.append("Chưa đặt tên đề.")
        if not self.data["du_an"]:
            problems.append("Đề chưa có dự án nào.")
        names = [p.get("file", "").casefold() for p in self.data["du_an"]]
        for i, p in enumerate(self.data["du_an"], start=1):
            where = f"Dự án {i}"
            if not p.get("file"):
                problems.append(f"{where}: chưa chọn file gốc.")
            elif names.count(p["file"].casefold()) > 1:
                problems.append(f"{where}: trùng tên file gốc '{p['file']}' với dự án khác.")
            if not p["nhiem_vu"]:
                problems.append(f"{where}: chưa có nhiệm vụ.")
            for j, t in enumerate(p["nhiem_vu"], start=1):
                if not t.get("yeu_cau"):
                    problems.append(f"{where} › nhiệm vụ {j}: chưa ghi yêu cầu.")
                if not t.get("cham"):
                    problems.append(f"{where} › nhiệm vụ {j}: chưa có luật chấm.")
        return problems

    def save(self, quiet=False) -> bool:
        self._commit_task()
        problems = self._validate()
        if problems:
            messagebox.showwarning("Chưa lưu được", "\n".join(problems[:12]), parent=self.win)
            return False
        self.data["ten"] = self.ten.get().strip()
        self.data["mon"] = self._mon_code()
        folder = self.folder or custom.new_exam_folder(self.data["ten"])
        try:
            folder.mkdir(parents=True, exist_ok=True)
            for p in self.data["du_an"]:
                if p.get("_src"):
                    if Path(p["_src"]).resolve() != (folder / p["file"]).resolve():
                        shutil.copyfile(p["_src"], folder / p["file"])
                if p.get("_ans"):
                    (folder / "dap_an").mkdir(exist_ok=True)
                    shutil.copyfile(p["_ans"], folder / "dap_an" / p["file"])
            clean = {**self.data, "du_an": [{k: v for k, v in p.items() if not k.startswith("_")}
                                            for p in self.data["du_an"]]}
            custom.save_exam_json(folder, clean)
        except OSError as exc:
            messagebox.showerror("Lỗi khi lưu", str(exc), parent=self.win)
            return False
        for p in self.data["du_an"]:
            p.pop("_src", None)
            p.pop("_ans", None)
        self.folder = folder
        self.dirty = False
        self._refresh_files()
        self.on_saved()
        if not quiet:
            messagebox.showinfo("Đã lưu", f"Đã lưu đề vào:\n{folder}", parent=self.win)
        return True

    def save_and_check(self):
        if self.save(quiet=True):
            CheckReport(self.win, self.folder)

    def close(self):
        if self.dirty and not messagebox.askyesno("Chưa lưu", "Đề có thay đổi chưa lưu. Đóng không lưu?",
                                                  parent=self.win):
            return
        self.win.destroy()


def describe_spec(spec: dict) -> str:
    name = spec.get("luat", "?")
    if name not in rules.RULES:
        return f"⚠ luật không tồn tại: {name}"
    params = {p["name"]: p for p in rules.rule_params(name)}
    parts = [f"{params[k]['label'].split(' (')[0]}: {rules.format_value(params[k]['kind'], v)}"
             for k, v in spec.items() if k != "luat" and k in params]
    return f"{rules.rule_title(name)}" + (f"  —  {'; '.join(parts)}" if parts else "")


class RuleDialog:
    """Chọn luật chấm và điền tham số (form sinh tự động theo luật)."""

    def __init__(self, parent, mon: str, spec: dict | None, on_ok):
        self.on_ok = on_ok
        self.names = rules.rules_for(mon)
        w = self.win = tk.Toplevel(parent, bg=BG, padx=22, pady=18)
        w.title("Luật chấm")
        w.minsize(560, 380)
        label(w, "Kiểm tra điều gì?", bold=True).pack(anchor="w")
        self.choice = tk.StringVar()
        self.titles = [rules.rule_title(n) for n in self.names]
        cb = ttk.Combobox(w, textvariable=self.choice, values=self.titles, state="readonly", width=70)
        cb.pack(fill="x", pady=(4, 2), ipady=2)
        cb.bind("<<ComboboxSelected>>", lambda e: self.build_form())
        self.doc = label(w, "", muted=True, size=9, wraplength=520)
        self.doc.pack(anchor="w", pady=(2, 8))
        self.form = tk.Frame(w, bg=BG)
        self.form.pack(fill="both", expand=True)
        self.form.columnconfigure(1, weight=1)
        bar = tk.Frame(w, bg=BG)
        bar.pack(fill="x", pady=(14, 0))
        Btn(bar, "OK", self.ok, kind="primary").pack(side="right")
        Btn(bar, "Hủy", w.destroy).pack(side="right", padx=8)
        self.vars: dict = {}
        start = spec.get("luat") if spec and spec.get("luat") in self.names else self.names[0]
        self.choice.set(self.titles[self.names.index(start)])
        self.build_form(spec if spec and spec.get("luat") == start else None)
        modal(w, parent)

    @property
    def name(self):
        return self.names[self.titles.index(self.choice.get())]

    def build_form(self, spec=None):
        for child in self.form.winfo_children():
            child.destroy()
        self.vars = {}
        fn = rules.RULES[self.name]
        params = rules.rule_params(self.name)
        doc = " ".join((fn.__doc__ or "").split())
        for p in params:   # thay tên biến trong mô tả bằng nhãn dễ hiểu
            doc = doc.replace(f"`{p['name']}`", "«" + p["label"].split(" (")[0].lower() + "»")
        self.doc.config(text=doc.replace("`", ""))
        for row, p in enumerate(params):
            text = p["label"] + ("" if p["required"] else "  (tuỳ chọn)")
            label(self.form, text, size=9, bold=p["required"]).grid(row=row, column=0, sticky="w", pady=4,
                                                                    padx=(0, 12))
            value = spec.get(p["name"]) if spec else None
            if value is None and not p["required"]:
                value = p["default"]
            var = tk.StringVar(value=rules.format_value(p["kind"], value))
            if p["kind"] in ("choice", "bool"):
                options = p["choices"] if p["kind"] == "choice" else list(rules.BOOL_TEXT)
                values = ([""] if not p["required"] and p["default"] is None else []) + options
                widget = ttk.Combobox(self.form, textvariable=var, values=values, state="readonly")
            else:
                widget = entry(self.form, var)
            widget.grid(row=row, column=1, sticky="ew", ipady=2)
            self.vars[p["name"]] = (var, p)

    def ok(self):
        spec = {"luat": self.name}
        try:
            for name, (var, p) in self.vars.items():
                value = rules.parse_value(p["kind"], var.get())
                if value is None:
                    if p["required"]:
                        raise ValueError(f"chưa điền “{p['label']}”")
                    continue
                if value != p["default"]:
                    spec[name] = value
        except ValueError as exc:
            messagebox.showerror("Thiếu thông tin", str(exc), parent=self.win)
            return
        err = rules.validate(spec)
        if err:
            messagebox.showerror("Lỗi", err, parent=self.win)
            return
        self.win.destroy()
        self.on_ok(spec)


class CheckReport:
    """Kết quả kiểm tra đề: file gốc phải SAI, file đáp án phải ĐÚNG."""

    def __init__(self, parent, folder: Path):
        errors, rows = custom.check_exam(folder)
        w = tk.Toplevel(parent, bg=BG, padx=20, pady=16)
        w.title("Kiểm tra đề")
        w.geometry("980x520")
        if errors:
            label(w, "Đề còn lỗi:", bold=True, size=12).pack(anchor="w")
            label(w, "\n".join("• " + e for e in errors), fg=DANGER).pack(anchor="w", pady=8)
            Btn(w, "Đóng", w.destroy).pack(anchor="e")
            modal(w, parent)
            return
        problems = sum(r["start_ok"] for r in rows) + sum(r["answer_ok"] is False for r in rows)
        missing = sum(r["answer_ok"] is None for r in rows)
        head = tk.Frame(w, bg=BG)
        head.pack(fill="x", pady=(0, 10))
        if problems:
            pill(head, f"CẦN XEM LẠI {problems} CHỖ", DANGER, "#fee2e2", size=11).pack(side="left")
        elif missing:
            pill(head, "CHƯA ĐỦ ĐÁP ÁN", WARN, WARN_BG, size=11).pack(side="left")
        else:
            pill(head, "ĐỀ ỔN", SUCCESS, "#dcfce7", size=11).pack(side="left")
        label(head, "   File gốc phải chấm SAI (chưa làm);  file đáp án phải chấm ĐÚNG.", muted=True
              ).pack(side="left")
        box, tree = make_tree(w, [("p", "Dự án", 180, False), ("t", "Nhiệm vụ", 460, True),
                                  ("s", "File gốc", 120, False), ("a", "File đáp án", 140, False)], height=12)
        box.pack(fill="both", expand=True)
        for r in rows:
            start = "ĐÚNG sẵn ⚠" if r["start_ok"] else "sai ✓"
            ans = "chưa có" if r["answer_ok"] is None else ("đúng ✓" if r["answer_ok"] else "SAI ✗")
            bad = r["start_ok"] or r["answer_ok"] is False
            tree.insert("", "end", values=(r["project"], f"{r['index']}. {r['task']}", start, ans),
                        tags=("bad",) if bad else ("muted",) if r["answer_ok"] is None else ("ok",))
        tips = []
        if any(r["start_ok"] for r in rows):
            tips.append("“ĐÚNG sẵn”: luật chấm quá dễ – file chưa làm đã đạt. Hãy siết điều kiện.")
        if any(r["answer_ok"] is False for r in rows):
            tips.append("“SAI” ở đáp án: làm đúng mà vẫn bị chấm sai – kiểm tra lại tham số của luật.")
        if missing:
            tips.append("Chọn “file đáp án” cho từng dự án để kiểm tra chiều đúng.")
        if tips:
            label(w, "\n".join("• " + t for t in tips), size=9, wraplength=920).pack(anchor="w", pady=(10, 0))
        Btn(w, "Đóng", w.destroy).pack(anchor="e", pady=(10, 0))
        modal(w, parent)
