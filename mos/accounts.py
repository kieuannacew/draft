"""Tài khoản người dùng (lưu trong DATA_DIR/tai_khoan.json).

Vai trò:
  * quan_tri – quản trị: tạo/khóa tài khoản, soạn đề, xem kết quả mọi người.
  * giao_vien – giáo viên (chế độ Lớp học trực tuyến): tạo lớp, xem kết quả học sinh lớp mình.
  * hoc_vien – chỉ làm bài và xem kết quả của mình.

Khi dùng máy chủ lớp học (mos/lop_hoc.py), tài khoản thật nằm trên máy chủ; máy chỉ giữ bản sao (mirror) để
đăng nhập được lúc mất mạng.

Lần chạy đầu tiên tự tạo tài khoản quản trị  admin / admin  và bắt đổi mật khẩu
khi đăng nhập. Mật khẩu không lưu dạng chữ mà lưu mã băm PBKDF2 + salt.
"""
from __future__ import annotations

import hashlib
import hmac
import json
import os
import re
from dataclasses import asdict, dataclass
from datetime import date, datetime
from pathlib import Path

from .core import DATA_DIR
from .i18n import tr

ADMIN, TEACHER, STUDENT = "quan_tri", "giao_vien", "hoc_vien"
ROLE_NAMES = {ADMIN: "Quản trị", TEACHER: "Giáo viên", STUDENT: "Học viên"}
DEFAULT_ADMIN = ("admin", "admin")
USERNAME_RE = re.compile(r"[a-z0-9_.]{3,32}")
MIN_PASSWORD = 4
_ITER = 200_000


@dataclass
class Account:
    username: str
    ho_ten: str
    vai_tro: str = STUDENT
    salt: str = ""
    hash: str = ""
    khoa: bool = False                 # bị khóa
    han_dung: str | None = None        # "YYYY-MM-DD" – hết hạn sau ngày này
    doi_mat_khau: bool = False         # bắt đổi mật khẩu ở lần đăng nhập tới
    tao_luc: str = ""
    nguon: str = ""                    # "may_chu" = bản sao tài khoản trên máy chủ lớp học

    @property
    def is_admin(self) -> bool:
        return self.vai_tro == ADMIN

    @property
    def is_teacher(self) -> bool:
        """Giáo viên hoặc quản trị: quản lý lớp học."""
        return self.vai_tro in (ADMIN, TEACHER)

    def expired(self, today: date | None = None) -> bool:
        return bool(self.han_dung) and (today or date.today()) > date.fromisoformat(self.han_dung)


def _hash(password: str, salt: bytes) -> str:
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, _ITER).hex()


class AccountError(ValueError):
    pass


class AccountStore:
    def __init__(self, path: Path | None = None):
        self.path = path or DATA_DIR / "tai_khoan.json"
        self.accounts: dict[str, Account] = {}
        self.load()
        if not self.accounts:
            name, pw = DEFAULT_ADMIN
            self.create(name, "Quản trị viên", pw, ADMIN, must_change=True)

    # ------------------------------------------------------------ lưu trữ
    def load(self) -> None:
        try:
            raw = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            raw = {}
        self.accounts = {k: Account(**v) for k, v in raw.items()}

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps({k: asdict(a) for k, a in self.accounts.items()},
                                  ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, self.path)

    # ------------------------------------------------------------ thao tác
    def list(self) -> list[Account]:
        return sorted(self.accounts.values(), key=lambda a: (a.vai_tro != ADMIN, a.username))

    def get(self, username: str) -> Account | None:
        return self.accounts.get(username.strip().lower())

    def create(self, username: str, ho_ten: str, password: str, vai_tro: str = STUDENT,
               han_dung: str | None = None, must_change: bool = True) -> Account:
        username = username.strip().lower()
        if not USERNAME_RE.fullmatch(username):
            raise AccountError("Tên đăng nhập 3–32 ký tự, chỉ gồm chữ thường không dấu, số, dấu _ hoặc .")
        if username in self.accounts:
            raise AccountError(tr("Tài khoản '{user}' đã tồn tại.").format(user=username))
        if vai_tro not in ROLE_NAMES:
            raise AccountError("Vai trò không hợp lệ.")
        acc = Account(username, ho_ten.strip() or username, vai_tro,
                      han_dung=_check_date(han_dung), doi_mat_khau=must_change,
                      tao_luc=datetime.now().strftime("%Y-%m-%d %H:%M"))
        self._set_pw(acc, password)
        self.accounts[username] = acc
        self.save()
        return acc

    def update(self, username: str, *, ho_ten=None, vai_tro=None, khoa=None, han_dung="giu") -> Account:
        acc = self._require(username)
        if vai_tro is not None and vai_tro not in ROLE_NAMES:
            raise AccountError("Vai trò không hợp lệ.")
        if acc.is_admin and ((vai_tro and vai_tro != ADMIN) or khoa) and self._admins() == 1:
            raise AccountError("Phải còn ít nhất một tài khoản quản trị đang hoạt động.")
        if ho_ten is not None:
            acc.ho_ten = ho_ten.strip() or acc.username
        if vai_tro is not None:
            acc.vai_tro = vai_tro
        if khoa is not None:
            acc.khoa = khoa
        if han_dung != "giu":
            acc.han_dung = _check_date(han_dung)
        self.save()
        return acc

    def set_password(self, username: str, password: str, must_change: bool = False) -> None:
        acc = self._require(username)
        self._set_pw(acc, password)
        acc.doi_mat_khau = must_change
        self.save()

    def delete(self, username: str) -> None:
        acc = self._require(username)
        if acc.is_admin and self._admins() == 1:
            raise AccountError("Không thể xóa tài khoản quản trị cuối cùng.")
        del self.accounts[acc.username]
        self.save()

    def mirror(self, username: str, ho_ten: str, vai_tro: str, password: str) -> Account:
        """Lưu bản sao tài khoản máy chủ (đăng nhập được khi mất mạng). Không kiểm tra độ dài mật khẩu."""
        username = username.strip().lower()
        acc = self.accounts.get(username)
        if acc is None:
            acc = Account(username, ho_ten or username, tao_luc=datetime.now().strftime("%Y-%m-%d %H:%M"))
            self.accounts[username] = acc
        acc.ho_ten = ho_ten or acc.ho_ten
        acc.vai_tro = vai_tro if vai_tro in ROLE_NAMES else STUDENT
        acc.khoa, acc.doi_mat_khau, acc.han_dung, acc.nguon = False, False, None, "may_chu"
        salt = os.urandom(16)
        acc.salt, acc.hash = salt.hex(), _hash(password, salt)
        self.save()
        return acc

    def authenticate(self, username: str, password: str) -> Account:
        acc = self.get(username)
        if acc is None or not hmac.compare_digest(acc.hash, _hash(password, bytes.fromhex(acc.salt))):
            raise AccountError("Sai tên đăng nhập hoặc mật khẩu.")
        if acc.khoa:
            raise AccountError("Tài khoản đã bị khóa. Liên hệ giáo viên.")
        if acc.expired():
            raise AccountError(tr("Tài khoản đã hết hạn ngày {date}. Liên hệ giáo viên.").format(date=acc.han_dung))
        return acc

    # ------------------------------------------------------------ nội bộ
    def _require(self, username: str) -> Account:
        acc = self.get(username)
        if acc is None:
            raise AccountError(tr("Không có tài khoản '{user}'.").format(user=username))
        return acc

    def _admins(self) -> int:
        return sum(a.is_admin and not a.khoa for a in self.accounts.values())

    @staticmethod
    def _set_pw(acc: Account, password: str) -> None:
        if len(password) < MIN_PASSWORD:
            raise AccountError(tr("Mật khẩu phải có ít nhất {n} ký tự.").format(n=MIN_PASSWORD))
        salt = os.urandom(16)
        acc.salt, acc.hash = salt.hex(), _hash(password, salt)


def _check_date(value: str | None) -> str | None:
    if not value:
        return None
    try:
        return date.fromisoformat(value.strip()).isoformat()
    except ValueError:
        raise AccountError("Hạn dùng phải có dạng YYYY-MM-DD, vd 2026-12-31.") from None
