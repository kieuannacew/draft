"""Supabase giả để kiểm thử phần Lớp học mà không cần tài khoản Supabase thật.

Gồm: Postgres thật (có schema `auth` rút gọn giống Supabase) + PostgREST thật + dịch vụ đăng nhập giả
(signup / token / user giống GoTrue), tất cả nằm sau một cổng như Supabase: /auth/v1/... và /rest/v1/...

Chỉ chạy khi có biến môi trường:
    MOS_TEST_PG      chuỗi kết nối Postgres quyền superuser, vd postgresql://postgres:postgres@localhost:5432/postgres
    POSTGREST_BIN    đường dẫn file chạy PostgREST (tải ở github.com/PostgREST/postgrest/releases)
"""
from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import secrets
import socket
import subprocess
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCHEMA = ROOT / "may_chu" / "supabase_lop_hoc.sql"
JWT_SECRET = "bi-mat-jwt-cho-kiem-thu-dai-hon-32-ky-tu"

STUB = r"""
do $$ begin
  if not exists (select 1 from pg_roles where rolname = 'anon') then create role anon nologin; end if;
  if not exists (select 1 from pg_roles where rolname = 'authenticated') then create role authenticated nologin; end if;
  if not exists (select 1 from pg_roles where rolname = 'authenticator') then
    create role authenticator login password 'authenticator' noinherit;
  end if;
end $$;
grant anon, authenticated to authenticator;
create schema if not exists auth;
create schema if not exists extensions;
create extension if not exists pgcrypto with schema extensions;
create table if not exists auth.users (
  id uuid primary key default gen_random_uuid(),
  email text unique not null,
  encrypted_password text not null,
  raw_user_meta_data jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
create or replace function auth.uid() returns uuid language sql stable as $f$
  select nullif(nullif(current_setting('request.jwt.claims', true), '')::json ->> 'sub', '')::uuid
$f$;
grant usage on schema auth, extensions, public to anon, authenticated;
grant execute on function auth.uid() to anon, authenticated;
"""


def available() -> bool:
    return bool(os.environ.get("MOS_TEST_PG") and os.environ.get("POSTGREST_BIN"))


def _free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _b64(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode()


def make_jwt(sub: str, seconds: int = 3600, role: str = "authenticated") -> str:
    head = _b64(json.dumps({"alg": "HS256", "typ": "JWT"}).encode())
    claims = {"role": role, "exp": int(time.time()) + seconds}
    if sub:
        claims["sub"] = sub
    body = _b64(json.dumps(claims).encode())
    sig = hmac.new(JWT_SECRET.encode(), f"{head}.{body}".encode(), hashlib.sha256).digest()
    return f"{head}.{body}.{_b64(sig)}"


ANON_KEY = make_jwt("", 10 * 365 * 86400, "anon")      # giống khóa "anon public" của Supabase
PUBLISHABLE_KEY = "sb_publishable_khoa_kieu_moi"        # khóa kiểu mới (không phải JWT)


class FakeSupabase:
    def __init__(self):
        admin_uri = os.environ["MOS_TEST_PG"]
        self.db = "mos_lop_hoc_test_" + secrets.token_hex(4)
        self.admin_uri = admin_uri
        parts = urllib.parse.urlsplit(admin_uri)
        self.uri = urllib.parse.urlunsplit(parts._replace(path="/" + self.db))
        self.rest_port, self.port = _free_port(), _free_port()
        self.url = f"http://127.0.0.1:{self.port}"
        self.refresh: dict[str, str] = {}
        self.token_seconds = 3600            # test hết hạn token: đặt số âm
        self.proc = None
        self.http = None

    # ------------------------------------------------------------ Postgres
    def psql(self, sql: str, uri: str | None = None) -> list[list[str]]:
        out = subprocess.run(["psql", uri or self.uri, "-v", "ON_ERROR_STOP=1", "-At", "-F", "\t", "-q", "-c", sql],
                             capture_output=True, text=True)
        if out.returncode:
            raise RuntimeError(out.stderr)
        return [line.split("\t") for line in out.stdout.splitlines() if line]

    def run_file(self, path: Path):
        out = subprocess.run(["psql", self.uri, "-v", "ON_ERROR_STOP=1", "-q", "-f", str(path)],
                             capture_output=True, text=True)
        if out.returncode:
            raise RuntimeError(out.stderr)

    def start(self) -> "FakeSupabase":
        self.psql(f"create database {self.db}", self.admin_uri)
        self.psql(STUB)
        self.run_file(SCHEMA)
        self.run_file(SCHEMA)            # chạy lại lần 2 vẫn phải được
        parts = urllib.parse.urlsplit(self.uri)
        rest_uri = urllib.parse.urlunsplit(parts._replace(
            netloc=f"authenticator:authenticator@{parts.hostname}:{parts.port or 5432}"))
        env = dict(os.environ, PGRST_DB_URI=rest_uri, PGRST_DB_SCHEMAS="public", PGRST_DB_ANON_ROLE="anon",
                   PGRST_JWT_SECRET=JWT_SECRET, PGRST_SERVER_PORT=str(self.rest_port),
                   PGRST_SERVER_HOST="127.0.0.1", PGRST_LOG_LEVEL="crit", PGRST_DB_MAX_ROWS="1000")
        self.proc = subprocess.Popen([os.environ["POSTGREST_BIN"]], env=env,
                                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        for _ in range(150):            # đợi PostgREST nạp xong cấu trúc CSDL
            try:
                urllib.request.urlopen(f"http://127.0.0.1:{self.rest_port}/", timeout=1)
                break
            except OSError:
                time.sleep(0.1)
        self.http = ThreadingHTTPServer(("127.0.0.1", self.port), _handler(self))
        threading.Thread(target=self.http.serve_forever, daemon=True).start()
        return self

    def stop(self):
        if self.http:
            self.http.shutdown()
        if self.proc:
            self.proc.terminate()
            self.proc.wait(10)
        self.psql(f"drop database if exists {self.db} with (force)", self.admin_uri)

    # ------------------------------------------------------------ auth giả
    def session(self, uid: str, email: str, meta: dict) -> dict:
        token = secrets.token_hex(16)
        self.refresh[token] = uid
        return {"access_token": make_jwt(uid, self.token_seconds), "token_type": "bearer",
                "expires_in": self.token_seconds, "refresh_token": token,
                "user": {"id": uid, "email": email, "user_metadata": meta}}


def _lit(value: str) -> str:
    return "'" + str(value).replace("'", "''") + "'"


def _handler(sb: FakeSupabase):
    class H(BaseHTTPRequestHandler):
        def log_message(self, *a):
            pass

        def _send(self, code: int, data=None, headers=None):
            raw = b"" if data is None else (data if isinstance(data, bytes) else json.dumps(data).encode())
            self.send_response(code)
            for k, v in (headers or {"Content-Type": "application/json"}).items():
                self.send_header(k, v)
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)

        def _body(self) -> bytes:
            n = int(self.headers.get("Content-Length") or 0)
            return self.rfile.read(n) if n else b""

        def _route(self, method: str):
            if self.headers.get("apikey") not in (ANON_KEY, PUBLISHABLE_KEY):
                return self._send(401, {"message": "No API key found in request"})
            path, _, query = self.path.partition("?")
            if path.startswith("/rest/v1/"):
                return self._proxy(method, path[len("/rest/v1"):], query)
            if path.startswith("/auth/v1/"):
                return self._auth(method, path[len("/auth/v1"):], urllib.parse.parse_qs(query))
            return self._send(404, {"message": "not found"})

        def _proxy(self, method, path, query):
            body = self._body() if method in ("POST", "PATCH", "PUT", "DELETE") else None
            req = urllib.request.Request(f"http://127.0.0.1:{sb.rest_port}{path}" + (f"?{query}" if query else ""),
                                         data=body, method=method)
            for k in ("Authorization", "Content-Type", "Prefer", "Accept", "Range", "Range-Unit"):
                if self.headers.get(k):
                    req.add_header(k, self.headers[k])
            if (self.headers.get("Authorization") or "").startswith("Bearer sb_"):
                return self._send(401, {"message": "Invalid JWT: publishable key sent as Authorization"})
            try:
                with urllib.request.urlopen(req, timeout=15) as r:
                    code, raw, hdr = r.status, r.read(), r.headers
            except urllib.error.HTTPError as e:
                code, raw, hdr = e.code, e.read(), e.headers
            out = {"Content-Type": hdr.get("Content-Type", "application/json")}
            if hdr.get("Content-Range"):
                out["Content-Range"] = hdr["Content-Range"]
            self._send(code, raw, out)

        def _auth(self, method, path, query):
            data = json.loads(self._body() or b"{}")
            if method == "GET" and path == "/settings":
                return self._send(200, {"disable_signup": False, "mailer_autoconfirm": True})
            if method == "POST" and path == "/signup":
                email, pw = data.get("email", "").lower(), data.get("password", "")
                if len(pw) < 6:
                    return self._send(422, {"code": 422, "error_code": "weak_password",
                                            "msg": "Password should be at least 6 characters."})
                if sb.psql(f"select 1 from auth.users where email = {_lit(email)}"):
                    return self._send(422, {"code": 422, "error_code": "user_already_exists",
                                            "msg": "User already registered"})
                meta = data.get("data") or {}
                uid = sb.psql("insert into auth.users (email, encrypted_password, raw_user_meta_data) values "
                              f"({_lit(email)}, extensions.crypt({_lit(pw)}, extensions.gen_salt('bf')), "
                              f"{_lit(json.dumps(meta))}::jsonb) returning id")[0][0]
                return self._send(200, sb.session(uid, email, meta))
            if method == "POST" and path == "/token":
                grant = (query.get("grant_type") or [""])[0]
                if grant == "password":
                    email, pw = data.get("email", "").lower(), data.get("password", "")
                    rows = sb.psql(f"select id, raw_user_meta_data from auth.users where email = {_lit(email)} "
                                   f"and encrypted_password = extensions.crypt({_lit(pw)}, encrypted_password)")
                    if not rows:
                        return self._send(400, {"code": 400, "error_code": "invalid_credentials",
                                                "msg": "Invalid login credentials"})
                    return self._send(200, sb.session(rows[0][0], email, json.loads(rows[0][1])))
                if grant == "refresh_token":
                    uid = sb.refresh.pop(data.get("refresh_token", ""), None)
                    if uid is None:
                        return self._send(400, {"code": 400, "error_code": "refresh_token_not_found",
                                                "msg": "Invalid Refresh Token: Refresh Token Not Found"})
                    return self._send(200, sb.session(uid, "", {}))
            if method == "PUT" and path == "/user":
                token = (self.headers.get("Authorization") or "").removeprefix("Bearer ")
                try:
                    head, body, sig = token.split(".")
                    good = _b64(hmac.new(JWT_SECRET.encode(), f"{head}.{body}".encode(), hashlib.sha256).digest())
                    claims = json.loads(base64.urlsafe_b64decode(body + "=="))
                    assert sig == good and claims["exp"] > time.time()
                except Exception:
                    return self._send(401, {"code": 401, "msg": "invalid JWT"})
                if len(data.get("password", "")) < 6:
                    return self._send(422, {"code": 422, "error_code": "weak_password",
                                            "msg": "Password should be at least 6 characters."})
                sb.psql(f"update auth.users set encrypted_password = extensions.crypt({_lit(data['password'])}, "
                        f"extensions.gen_salt('bf')) where id = {_lit(claims['sub'])}")
                return self._send(200, {"id": claims["sub"]})
            return self._send(404, {"msg": "not found"})

        def do_GET(self):
            self._route("GET")

        def do_POST(self):
            self._route("POST")

        def do_PATCH(self):
            self._route("PATCH")

        def do_PUT(self):
            self._route("PUT")

        def do_DELETE(self):
            self._route("DELETE")

    return H
