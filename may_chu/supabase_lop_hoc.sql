-- =====================================================================================================
-- Máy chủ "Lớp học" cho app Luyện thi MOS (Supabase).
-- Cách dùng: Supabase → SQL Editor → New query → dán TOÀN BỘ file này → Run. Chạy lại nhiều lần vẫn an toàn.
--
-- Bảng:   ho_so (hồ sơ + vai trò), lop (lớp, mã lớp), thanh_vien (học sinh trong lớp),
--         ket_qua (điểm mỗi lần làm bài), tien_do (tiến độ bài giảng)
-- Vai trò: quan_tri (người đăng ký ĐẦU TIÊN), giao_vien (quản trị cấp), hoc_vien (mặc định)
-- Bảo mật: Row Level Security – học sinh chỉ thấy dữ liệu của mình; giáo viên thấy học sinh trong lớp mình;
--          quản trị thấy tất cả. Các thao tác đặc biệt đi qua hàm (RPC) có kiểm tra quyền.
-- =====================================================================================================

create extension if not exists pgcrypto with schema extensions;

-- ----------------------------------------------------------------------------------------------- bảng
create table if not exists public.ho_so (
    id        uuid primary key references auth.users (id) on delete cascade,
    username  text not null unique check (username ~ '^[a-z0-9_.]{3,32}$'),
    ho_ten    text not null default '',
    vai_tro   text not null default 'hoc_vien' check (vai_tro in ('quan_tri', 'giao_vien', 'hoc_vien')),
    khoa      boolean not null default false,
    tao_luc   timestamptz not null default now()
);

create table if not exists public.lop (
    id         uuid primary key default gen_random_uuid(),
    ten        text not null check (length(trim(ten)) > 0),
    ma         text not null unique,
    giao_vien  uuid not null references public.ho_so (id) on delete cascade,
    tao_luc    timestamptz not null default now()
);

create table if not exists public.thanh_vien (
    lop      uuid not null references public.lop (id) on delete cascade,
    hoc_vien uuid not null references public.ho_so (id) on delete cascade,
    vao_luc  timestamptz not null default now(),
    primary key (lop, hoc_vien)
);

create table if not exists public.ket_qua (
    id        uuid primary key default gen_random_uuid(),
    hoc_vien  uuid not null default auth.uid() references public.ho_so (id) on delete cascade,
    ma_client text not null,                       -- mã do app tạo, để gửi lại không bị trùng
    exam      text not null default '',
    exam_en   text not null default '',
    code      text not null default '',            -- WORD / EXCEL / POWERPOINT
    mode      text not null default '',            -- training / chapter / testing
    score     integer not null default 0,
    correct   integer not null default 0,
    total     integer not null default 0,
    passed    boolean not null default false,
    duration  integer not null default 0,          -- giây
    wrong     jsonb not null default '[]'::jsonb,  -- các câu sai
    lam_luc   timestamptz not null default now(),
    unique (hoc_vien, ma_client)
);
create index if not exists ket_qua_hoc_vien_idx on public.ket_qua (hoc_vien, lam_luc desc);

create table if not exists public.tien_do (
    hoc_vien  uuid not null default auth.uid() references public.ho_so (id) on delete cascade,
    bai       text not null,                       -- id bài giảng
    ten_bai   text not null default '',
    xem       integer not null default 0,          -- slide xa nhất đã xem (video / SCORM: 1)
    tong      integer not null default 0,
    xong      boolean not null default false,
    diem      text not null default '',            -- điểm bài SCORM (vd "80/100")
    cap_nhat  timestamptz not null default now(),
    primary key (hoc_vien, bai)
);

-- ------------------------------------------------------------------------------------ hàm hỗ trợ quyền
create or replace function public.vai_tro_cua_toi() returns text
language sql stable security definer set search_path = public as $$
    select vai_tro from public.ho_so where id = auth.uid() and not khoa
$$;

create or replace function public.la_quan_tri() returns boolean
language sql stable security definer set search_path = public as $$
    select coalesce(public.vai_tro_cua_toi() = 'quan_tri', false)
$$;

create or replace function public.day_hoc_vien(hv uuid) returns boolean
language sql stable security definer set search_path = public as $$
    select exists (select 1 from public.thanh_vien t join public.lop l on l.id = t.lop
                   where t.hoc_vien = hv and l.giao_vien = auth.uid())
$$;

create or replace function public.la_chu_lop(lop_id uuid) returns boolean
language sql stable security definer set search_path = public as $$
    select exists (select 1 from public.lop where id = lop_id and giao_vien = auth.uid())
$$;

create or replace function public.trong_lop(lop_id uuid) returns boolean
language sql stable security definer set search_path = public as $$
    select exists (select 1 from public.thanh_vien where lop = lop_id and hoc_vien = auth.uid())
$$;

-- ------------------------------------------------------------------ tự tạo hồ sơ khi đăng ký tài khoản
create or replace function public.tao_ho_so() returns trigger
language plpgsql security definer set search_path = public as $$
begin
    insert into public.ho_so (id, username, ho_ten, vai_tro)
    values (new.id,
            lower(coalesce(new.raw_user_meta_data ->> 'username', split_part(new.email, '@', 1))),
            coalesce(nullif(trim(new.raw_user_meta_data ->> 'ho_ten'), ''),
                     new.raw_user_meta_data ->> 'username', split_part(new.email, '@', 1)),
            case when exists (select 1 from public.ho_so) then 'hoc_vien' else 'quan_tri' end);
    return new;
end $$;

drop trigger if exists tao_ho_so on auth.users;
create trigger tao_ho_so after insert on auth.users for each row execute function public.tao_ho_so();

-- --------------------------------------------------------------------------------- Row Level Security
alter table public.ho_so      enable row level security;
alter table public.lop        enable row level security;
alter table public.thanh_vien enable row level security;
alter table public.ket_qua    enable row level security;
alter table public.tien_do    enable row level security;

revoke all on public.ho_so, public.lop, public.thanh_vien, public.ket_qua, public.tien_do from anon, authenticated;
grant select on public.ho_so to authenticated;
grant select, delete on public.lop to authenticated;
grant update (ten) on public.lop to authenticated;
grant select, delete on public.thanh_vien to authenticated;
grant select, insert on public.ket_qua to authenticated;
grant delete on public.ket_qua to authenticated;
grant select on public.tien_do to authenticated;

drop policy if exists ho_so_xem on public.ho_so;
create policy ho_so_xem on public.ho_so for select to authenticated
    using (id = auth.uid() or public.la_quan_tri() or public.day_hoc_vien(id));

drop policy if exists lop_xem on public.lop;
create policy lop_xem on public.lop for select to authenticated
    using (giao_vien = auth.uid() or public.la_quan_tri() or public.trong_lop(id));
drop policy if exists lop_sua on public.lop;
create policy lop_sua on public.lop for update to authenticated
    using (giao_vien = auth.uid() or public.la_quan_tri()) with check (true);
drop policy if exists lop_xoa on public.lop;
create policy lop_xoa on public.lop for delete to authenticated
    using (giao_vien = auth.uid() or public.la_quan_tri());

drop policy if exists thanh_vien_xem on public.thanh_vien;
create policy thanh_vien_xem on public.thanh_vien for select to authenticated
    using (hoc_vien = auth.uid() or public.la_chu_lop(lop) or public.la_quan_tri());
drop policy if exists thanh_vien_xoa on public.thanh_vien;
create policy thanh_vien_xoa on public.thanh_vien for delete to authenticated
    using (hoc_vien = auth.uid() or public.la_chu_lop(lop) or public.la_quan_tri());

drop policy if exists ket_qua_xem on public.ket_qua;
create policy ket_qua_xem on public.ket_qua for select to authenticated
    using (hoc_vien = auth.uid() or public.la_quan_tri() or public.day_hoc_vien(hoc_vien));
drop policy if exists ket_qua_gui on public.ket_qua;
create policy ket_qua_gui on public.ket_qua for insert to authenticated
    with check (hoc_vien = auth.uid() and public.vai_tro_cua_toi() is not null);
drop policy if exists ket_qua_xoa on public.ket_qua;
create policy ket_qua_xoa on public.ket_qua for delete to authenticated
    using (public.la_quan_tri());

drop policy if exists tien_do_xem on public.tien_do;
create policy tien_do_xem on public.tien_do for select to authenticated
    using (hoc_vien = auth.uid() or public.la_quan_tri() or public.day_hoc_vien(hoc_vien));

-- ------------------------------------------------------------------------------------------- RPC
-- Học sinh vào lớp bằng mã lớp.
create or replace function public.vao_lop(ma_lop text) returns json
language plpgsql security definer set search_path = public as $$
declare l public.lop;
begin
    if public.vai_tro_cua_toi() is null then raise exception 'chua_dang_nhap'; end if;
    select * into l from public.lop where ma = upper(trim(ma_lop));
    if not found then raise exception 'sai_ma_lop'; end if;
    insert into public.thanh_vien (lop, hoc_vien) values (l.id, auth.uid()) on conflict do nothing;
    return json_build_object('id', l.id, 'ten', l.ten);
end $$;

-- Các lớp của tôi (học sinh: lớp đã vào; giáo viên: lớp mình dạy; quản trị: mọi lớp).
create or replace function public.lop_cua_toi() returns table (id uuid, ten text, ma text, giao_vien uuid,
                                                                ten_giao_vien text, si_so bigint, tao_luc timestamptz)
language sql stable security definer set search_path = public as $$
    select l.id, l.ten,
           case when l.giao_vien = auth.uid() or public.la_quan_tri() then l.ma else '' end,
           l.giao_vien, g.ho_ten,
           (select count(*) from public.thanh_vien t where t.lop = l.id), l.tao_luc
    from public.lop l join public.ho_so g on g.id = l.giao_vien
    where public.vai_tro_cua_toi() is not null
      and (l.giao_vien = auth.uid() or public.la_quan_tri() or public.trong_lop(l.id))
    order by l.tao_luc
$$;

-- Giáo viên / quản trị tạo lớp; mã lớp 6 ký tự do máy chủ sinh.
create or replace function public.tao_lop(ten_lop text) returns public.lop
language plpgsql security definer set search_path = public as $$
declare
    chu constant text := 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
    m text; l public.lop; i int;
begin
    if coalesce(public.vai_tro_cua_toi(), '') not in ('giao_vien', 'quan_tri') then
        raise exception 'khong_du_quyen';
    end if;
    loop
        m := '';
        for i in 1..6 loop
            m := m || substr(chu, 1 + floor(random() * length(chu))::int, 1);
        end loop;
        exit when not exists (select 1 from public.lop where ma = m);
    end loop;
    insert into public.lop (ten, ma, giao_vien) values (trim(ten_lop), m, auth.uid()) returning * into l;
    return l;
end $$;

-- Gửi tiến độ bài giảng: chỉ tăng, không giảm.
create or replace function public.luu_tien_do(p_bai text, p_ten_bai text, p_xem int, p_tong int,
                                              p_xong boolean, p_diem text default '') returns void
language plpgsql security definer set search_path = public as $$
begin
    if public.vai_tro_cua_toi() is null then raise exception 'chua_dang_nhap'; end if;
    insert into public.tien_do as t (hoc_vien, bai, ten_bai, xem, tong, xong, diem)
    values (auth.uid(), p_bai, coalesce(p_ten_bai, ''), greatest(p_xem, 0), greatest(p_tong, 0),
            coalesce(p_xong, false), coalesce(p_diem, ''))
    on conflict (hoc_vien, bai) do update set
        ten_bai  = excluded.ten_bai,
        xem      = greatest(t.xem, excluded.xem),
        tong     = excluded.tong,
        xong     = t.xong or excluded.xong,
        diem     = coalesce(nullif(excluded.diem, ''), t.diem),
        cap_nhat = now();
end $$;

-- Tự đổi họ tên.
create or replace function public.doi_ho_ten(p_ho_ten text) returns void
language sql security definer set search_path = public as $$
    update public.ho_so set ho_ten = trim(p_ho_ten) where id = auth.uid() and length(trim(p_ho_ten)) > 0
$$;

-- Quản trị đổi vai trò (không được bỏ quản trị cuối cùng).
create or replace function public.dat_vai_tro(p_id uuid, p_vai_tro text) returns void
language plpgsql security definer set search_path = public as $$
begin
    if not public.la_quan_tri() then raise exception 'khong_du_quyen'; end if;
    if p_vai_tro not in ('quan_tri', 'giao_vien', 'hoc_vien') then raise exception 'vai_tro_sai'; end if;
    if p_vai_tro <> 'quan_tri' and (select vai_tro from public.ho_so where id = p_id) = 'quan_tri'
       and (select count(*) from public.ho_so where vai_tro = 'quan_tri' and not khoa) <= 1 then
        raise exception 'quan_tri_cuoi';
    end if;
    update public.ho_so set vai_tro = p_vai_tro where id = p_id;
end $$;

-- Quản trị khóa / mở khóa tài khoản.
create or replace function public.dat_khoa(p_id uuid, p_khoa boolean) returns void
language plpgsql security definer set search_path = public as $$
begin
    if not public.la_quan_tri() then raise exception 'khong_du_quyen'; end if;
    if p_id = auth.uid() and p_khoa then raise exception 'khong_tu_khoa'; end if;
    update public.ho_so set khoa = p_khoa where id = p_id;
end $$;

-- Đặt lại mật khẩu: quản trị cho mọi người, giáo viên cho học sinh lớp mình.
create or replace function public.dat_lai_mat_khau(p_id uuid, p_mat_khau text) returns void
language plpgsql security definer set search_path = public, extensions as $$
begin
    if not (public.la_quan_tri() or (public.vai_tro_cua_toi() = 'giao_vien' and public.day_hoc_vien(p_id)
                                     and (select vai_tro from public.ho_so where id = p_id) = 'hoc_vien')) then
        raise exception 'khong_du_quyen';
    end if;
    if length(coalesce(p_mat_khau, '')) < 6 then raise exception 'mat_khau_ngan'; end if;
    update auth.users set encrypted_password = extensions.crypt(p_mat_khau, extensions.gen_salt('bf')),
                          updated_at = now()
    where id = p_id;
end $$;

-- Quản trị xóa tài khoản (không tự xóa mình).
create or replace function public.xoa_tai_khoan(p_id uuid) returns void
language plpgsql security definer set search_path = public as $$
begin
    if not public.la_quan_tri() then raise exception 'khong_du_quyen'; end if;
    if p_id = auth.uid() then raise exception 'khong_tu_xoa'; end if;
    delete from auth.users where id = p_id;
end $$;

-- Chỉ tài khoản đã đăng nhập mới gọi được các hàm trên.
do $$
declare f text;
begin
    foreach f in array array[
        'vai_tro_cua_toi()', 'la_quan_tri()', 'day_hoc_vien(uuid)', 'la_chu_lop(uuid)', 'trong_lop(uuid)',
        'vao_lop(text)', 'lop_cua_toi()', 'tao_lop(text)', 'luu_tien_do(text,text,int,int,boolean,text)',
        'doi_ho_ten(text)', 'dat_vai_tro(uuid,text)', 'dat_khoa(uuid,boolean)', 'dat_lai_mat_khau(uuid,text)',
        'xoa_tai_khoan(uuid)']
    loop
        execute format('revoke all on function public.%s from public, anon', f);
        execute format('grant execute on function public.%s to authenticated', f);
    end loop;
    execute 'revoke all on function public.tao_ho_so() from public, anon, authenticated';
end $$;

-- Báo PostgREST đọc lại cấu trúc mới.
notify pgrst, 'reload schema';
