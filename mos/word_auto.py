"""Chấm tự động các dạng câu Word của bộ đề nhập từ ngoài (bộ sinh đề MO-110 / Expert).

Mỗi câu trong bộ đề có mã dạng (vd "pic_size") và đề bài tiếng Anh. Hàm chấm của
từng dạng đọc thông số ngay trong đề bài (chữ trong ngoặc “…”, số đo 2.5", tên
style…) rồi mở file .docx học viên đã lưu để kiểm tra.

    grade(path, dang, de_bai, goc)  -> True / False, hoặc None nếu chưa hỗ trợ dạng này
    snapshot(path, dang, de_bai)    -> dict thông tin file gốc cần để chấm (lưu vào de.json)

Nguyên tắc: chấm lỏng phần cốt lõi của yêu cầu, chịu được cách Word thật lưu file
(style id khác nhau, run bị tách, mc:AlternateContent…).
"""
from __future__ import annotations

import hashlib
import os
import re
import zipfile
from datetime import datetime
from pathlib import Path

from lxml import etree

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "v": "urn:schemas-microsoft-com:vml",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "cs": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
}
W = "{%s}" % NS["w"]
P, TBL, T, TAB = W + "p", W + "tbl", W + "t", W + "tab"
FALLBACK = "{%s}Fallback" % NS["mc"]
TXBX = W + "txbxContent"
EMU_IN, TW_IN = 914400, 1440

GRADERS: dict = {}
SNAPS: dict = {}


def grader(*names):
    def deco(fn):
        for n in names:
            GRADERS[n] = fn
        return fn
    return deco


def snap(*names):
    def deco(fn):
        for n in names:
            SNAPS[n] = fn
        return fn
    return deco


# ====================================================================== đọc file


def norm(s) -> str:
    s = str(s or "").replace("\u00a0", " ").replace("’", "'").replace("‘", "'")
    return " ".join(s.split()).casefold()


def _nospace(s) -> str:
    return re.sub(r"\s+", "", norm(s))


def _attr(el, name, ns="w"):
    if el is None:
        return None
    return el.get("{%s}%s" % (NS[ns], name)) if ns else el.get(name)


def _on(el, attr="val") -> bool:
    """Thuộc tính bật/tắt kiểu <w:b/>, <w:b w:val="0"/>."""
    if el is None:
        return False
    return (_attr(el, attr) or "true").lower() not in ("0", "false", "off", "none")


class Doc:
    """File .docx đã giải nén, đọc XML theo nhu cầu."""

    def __init__(self, path):
        self.path = Path(path)
        with zipfile.ZipFile(self.path) as z:
            self.files = {n: z.read(n) for n in z.namelist()}
        self._xml: dict = {}
        self._styles = None

    def raw(self, name) -> str:
        data = self.files.get(name)
        return data.decode("utf-8", "replace") if data else ""

    def xml(self, name):
        if name not in self._xml:
            data = self.files.get(name)
            self._xml[name] = etree.fromstring(data) if data else None
        return self._xml[name]

    def glob(self, pattern) -> list[str]:
        rx = re.compile(pattern)
        return sorted(n for n in self.files if rx.fullmatch(n))

    @property
    def doc(self):
        return self.xml("word/document.xml")

    @property
    def body(self):
        return self.doc.find(W + "body")

    # ---------- quan hệ (rels)
    def rels(self, part: str) -> dict:
        folder, _, name = part.rpartition("/")
        root = self.xml(f"{folder}/_rels/{name}.rels")
        out = {}
        if root is None:
            return out
        for rel in root:
            target = rel.get("Target", "")
            if rel.get("TargetMode") == "External":
                out[rel.get("Id")] = target
            else:
                out[rel.get("Id")] = _join(folder, target)
        return out

    # ---------- style
    def styles(self) -> dict:
        """styleId -> phần tử <w:style>."""
        if self._styles is None:
            root = self.xml("word/styles.xml")
            self._styles = {} if root is None else {s.get(W + "styleId"): s for s in root.iter(W + "style")}
        return self._styles

    def style_name(self, sid) -> str:
        st = self.styles().get(sid)
        if st is None:
            return norm(sid)
        name = st.find(W + "name")
        return norm(_attr(name, "val") or sid)

    def style_by_name(self, name):
        for st in self.styles().values():
            if self.style_name(st.get(W + "styleId")) == norm(name):
                return st
        return None

    def default_style(self, kind="paragraph"):
        for st in self.styles().values():
            if st.get(W + "type") == kind and _on(st, "default"):
                return st
        return None

    def style_chain(self, sid, kind="paragraph"):
        seen, out = set(), []
        st = self.styles().get(sid) if sid else self.default_style(kind)
        while st is not None and id(st) not in seen:
            seen.add(id(st))
            out.append(st)
            based = st.find(W + "basedOn")
            st = self.styles().get(_attr(based, "val")) if based is not None else None
        return out


def _join(folder, target):
    if target.startswith("/"):
        return target[1:]
    parts = (folder.split("/") if folder else []) + target.split("/")
    out = []
    for p in parts:
        if p == "..":
            if out:
                out.pop()
        elif p and p != ".":
            out.append(p)
    return "/".join(out)


def _walk(el):
    """Duyệt con cháu, bỏ qua đoạn văn lồng, text box và phần mc:Fallback (bản sao VML)."""
    for ch in el:
        if not isinstance(ch.tag, str) or ch.tag in (FALLBACK, P, TXBX):
            continue
        yield ch
        yield from _walk(ch)


def ptext(p) -> str:
    out = []
    for el in _walk(p):
        if el.tag == T and el.text:
            out.append(el.text)
        elif el.tag == TAB and el.getparent().tag == W + "r":
            out.append("\t")
    return "".join(out)


def all_paras(root):
    """Mọi đoạn văn (kể cả trong bảng), không lấy trong text box / Fallback."""
    out = []

    def rec(el):
        for ch in el:
            if not isinstance(ch.tag, str) or ch.tag in (FALLBACK, TXBX):
                continue
            if ch.tag == P:
                out.append(ch)
            rec(ch)
    rec(root)
    return out


def blocks(d: Doc) -> list:
    """Các khối cấp thân bài theo thứ tự: đoạn văn và bảng (mở sdt / customXml)."""
    out = []

    def rec(el):
        for ch in el:
            if ch.tag in (P, TBL):
                out.append(ch)
            elif ch.tag in (W + "sdt", W + "sdtContent", W + "customXml", W + "smartTag"):
                rec(ch)
    rec(d.body)
    return out


def pstyle(d: Doc, p) -> str:
    ps = p.find(f"{W}pPr/{W}pStyle")
    sid = _attr(ps, "val")
    if sid is None:
        st = d.default_style()
        sid = st.get(W + "styleId") if st is not None else "Normal"
    return d.style_name(sid)


def is_h1(d: Doc, p) -> bool:
    name = pstyle(d, p)
    if name in ("heading 1",):
        return True
    lvl = p.find(f"{W}pPr/{W}outlineLvl")
    return _attr(lvl, "val") == "0"


def find_para(d: Doc, text, start=False, root=None):
    """Đoạn có chữ đúng bằng / bắt đầu bằng `text` (bỏ dấu … ở cuối)."""
    want = norm(str(text).rstrip(" …").rstrip("."))
    for p in all_paras(root if root is not None else d.body):
        t = norm(ptext(p))
        if not t:
            continue
        if (t.startswith(want) if start else (t == want or t.rstrip(".") == want)):
            return p
    return None


def heading_para(d: Doc, heading):
    want = norm(heading)
    bl = blocks(d)
    for b in bl:
        if b.tag == P and norm(ptext(b)) == want:
            return b
    for b in bl:
        if b.tag == P and norm(ptext(b)).startswith(want):
            return b
    return None


def section(d: Doc, heading) -> list:
    """Các khối sau tiêu đề `heading` cho tới tiêu đề cấp 1 tiếp theo."""
    bl = blocks(d)
    h = heading_para(d, heading)
    if h is None:
        return []
    i = bl.index(h)
    out = []
    for b in bl[i + 1:]:
        if b.tag == P and is_h1(d, b) and ptext(b).strip():
            break
        out.append(b)
    return out


def sec_paras(d, heading, nonempty=True):
    """Các đoạn văn thường (cấp thân bài, có chữ) trong section."""
    return [b for b in section(d, heading) if b.tag == P and (ptext(b).strip() or not nonempty)]


def sec_tables(d, heading):
    return [b for b in section(d, heading) if b.tag == TBL]


def quotes(text) -> list[str]:
    return [q.strip() for q in re.findall(r"[“\"]([^”\"]+)[”\"]", text)]


def sec_name(text):
    m = re.search(r"in the [“\"]([^”\"]+)[”\"] section", text, re.I) or \
        re.search(r"[“\"]([^”\"]+)[”\"] section", text, re.I)
    return m.group(1) if m else None


def inches(text) -> list[float]:
    return [float(x) for x in re.findall(r"(\d+(?:\.\d+)?)\"", text)]


def _near(a, b, tol):
    return a is not None and b is not None and abs(float(a) - float(b)) <= tol


def _int(x, default=None):
    try:
        return int(float(x))
    except (TypeError, ValueError):
        return default


# ---------- thuộc tính đoạn / run có tính kế thừa style


def ppr_find(d: Doc, p, path):
    """Tìm phần tử trong pPr của đoạn, rồi style, rồi docDefaults."""
    ppr = p.find(W + "pPr")
    if ppr is not None:
        el = ppr.find(path, NS)
        if el is not None:
            return el
    sid = _attr(p.find(f"{W}pPr/{W}pStyle"), "val")
    for st in d.style_chain(sid):
        el = st.find("w:pPr/" + path, NS)
        if el is not None:
            return el
    styles = d.xml("word/styles.xml")
    if styles is not None:
        return styles.find("w:docDefaults/w:pPrDefault/w:pPr/" + path, NS)
    return None


def ppr_attr(d, p, path, attr):
    """Giá trị thuộc tính (vd spacing/@before), lần theo kế thừa từng thuộc tính."""
    cands = []
    ppr = p.find(W + "pPr")
    if ppr is not None:
        cands.append(ppr)
    sid = _attr(p.find(f"{W}pPr/{W}pStyle"), "val")
    cands += [st.find(W + "pPr") for st in d.style_chain(sid)]
    styles = d.xml("word/styles.xml")
    if styles is not None:
        cands.append(styles.find("w:docDefaults/w:pPrDefault/w:pPr", NS))
    for c in cands:
        if c is None:
            continue
        el = c.find(path, NS)
        if el is not None and _attr(el, attr) is not None:
            return _attr(el, attr)
    return None


def rpr_find(d: Doc, run, p, name):
    """Phần tử rPr/name của run: trực tiếp → style ký tự → style đoạn → mặc định."""
    rpr = run.find(W + "rPr")
    if rpr is not None and rpr.find(W + name) is not None:
        return rpr.find(W + name)
    rs = _attr(run.find(f"{W}rPr/{W}rStyle"), "val")
    if rs:
        for st in d.style_chain(rs, "character"):
            el = st.find(f"{W}rPr/{W}{name}")
            if el is not None:
                return el
    sid = _attr(p.find(f"{W}pPr/{W}pStyle"), "val") if p is not None else None
    for st in d.style_chain(sid):
        el = st.find(f"{W}rPr/{W}{name}")
        if el is not None:
            return el
    styles = d.xml("word/styles.xml")
    if styles is not None:
        return styles.find(f"w:docDefaults/w:rPrDefault/w:rPr/w:{name}", NS)
    return None


def text_runs(p) -> list:
    return [r for r in _walk(p) if r.tag == W + "r" and any(t.tag == T and t.text for t in r)]


def runs_covering(p, text):
    """Các run chứa phần chữ `text` trong đoạn p (None nếu không thấy)."""
    runs = text_runs(p)
    full, spans, pos = "", [], 0
    for r in runs:
        t = "".join(x.text or "" for x in r if x.tag == T)
        spans.append((pos, pos + len(t), r))
        full += t
        pos += len(t)
    want = str(text).rstrip(" …")
    i = full.find(want)
    if i < 0:
        i = full.casefold().find(want.casefold())
    if i < 0:
        return None
    j = i + len(want)
    return [r for a, b, r in spans if a < j and b > i and full[max(a, i):min(b, j)].strip()]


# ---------- màu


STD_COLORS = {"dark red": "C00000", "red": "FF0000", "orange": "FFC000", "yellow": "FFFF00",
              "light green": "92D050", "green": "00B050", "light blue": "00B0F0", "blue": "0070C0",
              "dark blue": "002060", "purple": "7030A0", "black": "000000", "white": "FFFFFF"}
THEME_WORD = {"text 1": "text1", "text 2": "text2", "background 1": "background1", "background 2": "background2"}
TINT = {"80": "33", "60": "66", "40": "99", "25": "BF", "50": "80"}


def color_spec(desc: str) -> dict:
    """"Blue, Accent 1, Lighter 80%" -> {theme: accent1, tint: 33}; "Purple" -> {hex: 7030A0}."""
    desc = desc.strip().rstrip(".")
    low = desc.casefold()
    if low in ("automatic", "auto"):
        return {"auto": True}
    m = re.search(r"accent\s*(\d)", low)
    theme = f"accent{m.group(1)}" if m else next((v for k, v in THEME_WORD.items() if k in low), None)
    spec = {}
    if theme:
        spec["theme"] = theme
        m = re.search(r"lighter\s*(\d+)%", low)
        if m:
            spec["tint"] = TINT.get(m.group(1))
        m = re.search(r"darker\s*(\d+)%", low)
        if m:
            spec["shade"] = TINT.get(m.group(1))
        return spec
    key = low.split("(")[0].strip()
    if key in STD_COLORS:
        return {"hex": STD_COLORS[key]}
    return {}


def color_ok(el, spec, kind="color") -> bool:
    """So màu trên phần tử w:color / w:shd (kind="fill") / viền (kind="border")."""
    if el is None or not spec:
        return bool(spec.get("auto")) if spec else False
    if kind == "fill":
        val, th, tint, shade = (_attr(el, "fill"), _attr(el, "themeFill"), _attr(el, "themeFillTint"),
                                _attr(el, "themeFillShade"))
    elif kind == "border":
        val, th, tint, shade = (_attr(el, "color"), _attr(el, "themeColor"), _attr(el, "themeTint"),
                                _attr(el, "themeShade"))
    else:
        val, th, tint, shade = (_attr(el, "val"), _attr(el, "themeColor"), _attr(el, "themeTint"),
                                _attr(el, "themeShade"))
    if spec.get("auto"):
        return (val or "auto").lower() == "auto" and not th
    if "theme" in spec:
        if (th or "").casefold() != spec["theme"]:
            return False
        if spec.get("tint") and (tint or "").upper() != spec["tint"]:
            return False
        if spec.get("shade") and (shade or "").upper() != spec["shade"]:
            return False
        if not spec.get("tint") and not spec.get("shade") and (tint or shade):
            return False
        return True
    return (val or "").upper() == spec["hex"]


DML_THEME = {"text1": "tx1", "text2": "tx2", "background1": "bg1", "background2": "bg2"}


def dml_color_ok(el, spec) -> bool:
    """Màu DrawingML (a:solidFill …) trong phần tử el."""
    if el is None or not spec:
        return False
    if "theme" in spec:
        names = {spec["theme"], DML_THEME.get(spec["theme"], spec["theme"])}
        if spec["theme"] == "background1":
            names |= {"lt1", "bg1"}
        if spec["theme"] == "text1":
            names |= {"dk1", "tx1"}
        if any(c.get("val") in names for c in el.iter("{%s}schemeClr" % NS["a"])):
            return True
        return spec["theme"] == "background1" and any(
            (c.get("val") or "").upper() == "FFFFFF" for c in el.iter("{%s}srgbClr" % NS["a"]))
    if "hex" in spec:
        return any((c.get("val") or "").upper() == spec["hex"] for c in el.iter("{%s}srgbClr" % NS["a"]))
    return False


# ---------- đối tượng vẽ (ảnh, biểu đồ, SmartArt, text box)

URI = {"pic": "drawingml/2006/picture", "chart": "drawingml/2006/chart", "dgm": "drawingml/2006/diagram",
       "wps": "wordprocessingShape", "model3d": "model3d", "wpg": "wordprocessingGroup"}


def drawings(els, kind=None) -> list:
    """Các wp:inline / wp:anchor trong danh sách phần tử (không lấy bản Fallback)."""
    out = []
    for el in els:
        cands = [el] if el.tag in ("{%s}inline" % NS["wp"], "{%s}anchor" % NS["wp"]) else []
        cands += [x for x in el.iter("{%s}inline" % NS["wp"], "{%s}anchor" % NS["wp"])]
        for x in cands:
            if any(a.tag == FALLBACK for a in x.iterancestors()):
                continue
            gd = x.find(".//a:graphicData", NS)
            uri = gd.get("uri", "") if gd is not None else ""
            if kind is None or URI[kind] in uri or (kind == "pic" and gd is not None and gd.find(
                    ".//pic:pic", NS) is not None and "model3d" not in uri):
                if x not in out:
                    out.append(x)
    return out


def docpr(dr):
    return dr.find("wp:docPr", NS)


def extent(dr):
    e = dr.find("wp:extent", NS)
    return (_int(e.get("cx")), _int(e.get("cy"))) if e is not None else (None, None)


def chart_parts(d: Doc, drs) -> list[str]:
    rels = d.rels("word/document.xml")
    out = []
    for dr in drs:
        ch = dr.find(".//c:chart", NS)
        if ch is not None:
            target = rels.get(ch.get("{%s}id" % NS["r"]))
            if target and target in d.files:
                out.append(target)
    return out


def dgm_parts(d: Doc, dr) -> dict:
    rels = d.rels("word/document.xml")
    ids = dr.find(".//dgm:relIds", NS)
    if ids is None:
        return {}
    return {k: rels.get(ids.get("{%s}%s" % (NS["r"], k))) for k in ("dm", "lo", "qs", "cs")}


def sec_objects(d, text, kind):
    name = sec_name(text)
    els = section(d, name) if name else [d.body]
    return drawings(els, kind)


def txbx_text(dr) -> str:
    tb = dr.find(".//w:txbxContent", NS)
    return "\n".join(ptext(p) for p in tb.iter(P)) if tb is not None else ""


def textboxes(d: Doc) -> list:
    return [dr for dr in drawings([d.body], "wps") if dr.find(".//w:txbxContent", NS) is not None]


def textbox_starting(d, start):
    for dr in textboxes(d):
        if norm(txbx_text(dr)).startswith(norm(start)):
            return dr
    return None


# ---------- đánh số


def numbering(d: Doc, p) -> dict | None:
    """Thông tin danh sách của đoạn: numId, ilvl, numFmt, lvlText, start, font."""
    num_id = ilvl = None
    npr = p.find(f"{W}pPr/{W}numPr")
    if npr is not None:
        num_id = _attr(npr.find(W + "numId"), "val")
        ilvl = _attr(npr.find(W + "ilvl"), "val")
    if num_id is None:
        sid = _attr(p.find(f"{W}pPr/{W}pStyle"), "val")
        for st in d.style_chain(sid):
            snpr = st.find(f"{W}pPr/{W}numPr")
            if snpr is not None and snpr.find(W + "numId") is not None:
                num_id = _attr(snpr.find(W + "numId"), "val")
                ilvl = ilvl or _attr(snpr.find(W + "ilvl"), "val")
                break
    if num_id in (None, "0"):
        return None
    ilvl = ilvl or "0"
    root = d.xml("word/numbering.xml")
    if root is None:
        return None
    num = next((n for n in root.findall("w:num", NS) if _attr(n, "numId") == num_id), None)
    if num is None:
        return None
    abs_id = _attr(num.find(W + "abstractNumId"), "val")
    absn = next((a for a in root.findall("w:abstractNum", NS) if _attr(a, "abstractNumId") == abs_id), None)
    lvl = None
    start_override = None
    for ov in num.findall("w:lvlOverride", NS):
        if _attr(ov, "ilvl") == ilvl:
            so = ov.find(W + "startOverride")
            if so is not None:
                start_override = _attr(so, "val")
            if ov.find(W + "lvl") is not None:
                lvl = ov.find(W + "lvl")
    if lvl is None and absn is not None:
        lvl = next((x for x in absn.findall("w:lvl", NS) if _attr(x, "ilvl") == ilvl), None)
    info = {"numId": num_id, "ilvl": int(ilvl), "abs": abs_id, "startOverride": start_override,
            "numFmt": None, "lvlText": None, "start": None, "font": None}
    if lvl is not None:
        info["numFmt"] = _attr(lvl.find(W + "numFmt"), "val")
        info["lvlText"] = _attr(lvl.find(W + "lvlText"), "val")
        info["start"] = _attr(lvl.find(W + "start"), "val")
        f = lvl.find(f"{W}rPr/{W}rFonts")
        info["font"] = _attr(f, "ascii") or _attr(f, "hAnsi") or _attr(f, "cs")
    return info


def fmt_spec(fmt: str) -> tuple[str, str]:
    """"i." -> (lowerRoman, "%1."); "(1)" -> (decimal, "(%1)"); "A." -> (upperLetter, "%1.")."""
    fmt = fmt.strip()
    m = re.search(r"(ii?i?|II?I?|\d+|[a-zA-Z])", fmt)
    core = m.group(1) if m else "1"
    if core.isdigit():
        kind = "decimal"
    elif re.fullmatch(r"i+", core):
        kind = "lowerRoman"
    elif re.fullmatch(r"I+", core):
        kind = "upperRoman"
    elif core.islower():
        kind = "lowerLetter"
    else:
        kind = "upperLetter"
    return kind, fmt[:m.start()] + "%1" + fmt[m.end():] if m else "%1."


def paras_from(d, heading, start, count):
    ps = sec_paras(d, heading) if heading else [p for p in blocks(d) if p.tag == P and ptext(p).strip()]
    for i, p in enumerate(ps):
        if norm(ptext(p)).startswith(norm(start)):
            return ps[i:i + count]
    return []


# ---------- tiện ích khác


def fields(root) -> list[str]:
    """Mã các trường (field) trong phần tử: instrText ghép lại + fldSimple."""
    out = []
    for el in root.iter(W + "fldSimple"):
        out.append(el.get(W + "instr") or "")
    buf, depth = [], 0
    for el in root.iter(W + "fldChar", W + "instrText"):
        if el.tag == W + "fldChar":
            t = _attr(el, "fldCharType")
            if t == "begin":
                if depth == 0:
                    buf = []
                depth += 1
            elif t == "separate" and depth == 1:
                out.append("".join(buf))
            elif t == "end":
                if depth == 1 and buf and ("".join(buf) not in out[-1:]):
                    out.append("".join(buf))
                depth = max(0, depth - 1)
        elif depth >= 1:
            buf.append(el.text or "")
    return [" ".join(f.split()) for f in out if f.strip()]


def core_prop(d: Doc, name) -> str:
    root = d.xml("docProps/core.xml")
    if root is None:
        return ""
    for el in root:
        if etree.QName(el).localname == name:
            return el.text or ""
    return ""


def app_prop(d: Doc, name) -> str:
    root = d.xml("docProps/app.xml")
    if root is None:
        return ""
    el = root.find(f"ep:{name}", NS)
    return (el.text or "") if el is not None else ""


def settings(d: Doc):
    return d.xml("word/settings.xml")


def sect_prs(d: Doc) -> list:
    return list(d.body.iter(W + "sectPr"))


def final_sect(d: Doc):
    return d.body.find(W + "sectPr")


def doc_text(d: Doc) -> str:
    return "\n".join(ptext(p) for p in all_paras(d.body))


def md5(data: bytes) -> str:
    return hashlib.md5(data).hexdigest()


def pic_image(d: Doc, dr) -> bytes:
    blip = dr.find(".//a:blip", NS)
    if blip is None:
        return b""
    target = d.rels("word/document.xml").get(blip.get("{%s}embed" % NS["r"]))
    return d.files.get(target, b"")


def comment_ranges(d: Doc) -> dict:
    """{id comment: chữ được gắn comment} theo thứ tự tài liệu."""
    active, out = {}, {}
    for el in d.body.iter():
        if el.tag == W + "commentRangeStart":
            active[_attr(el, "id")] = []
        elif el.tag == W + "commentRangeEnd":
            cid = _attr(el, "id")
            if cid in active:
                out[cid] = "".join(active.pop(cid))
        elif el.tag == T and el.text and not any(a.tag == FALLBACK for a in el.iterancestors()):
            for buf in active.values():
                buf.append(el.text)
        elif el.tag == W + "commentReference":
            cid = _attr(el, "id")
            out.setdefault(cid, "")
    return out


def comments(d: Doc) -> dict:
    root = d.xml("word/comments.xml")
    if root is None:
        return {}
    return {_attr(c, "id"): "\n".join(ptext(p) for p in c.iter(P)) for c in root.findall("w:comment", NS)}


def revisions(d: Doc) -> dict:
    ins = ["".join(t.text or "" for t in e.iter(T)) for e in d.body.iter(W + "ins")]
    dele = ["".join(t.text or "" for t in e.iter(W + "delText")) for e in d.body.iter(W + "del")]
    fmt = len(list(d.body.iter(W + "rPrChange"))) + len(list(d.body.iter(W + "pPrChange")))
    return {"ins": [x for x in ins if x.strip()], "del": [x for x in dele if x.strip()], "fmt": fmt}


def title_para(d: Doc):
    for p in all_paras(d.body):
        if pstyle(d, p) == "title" and ptext(p).strip():
            return p
    for b in blocks(d):
        if b.tag == P and ptext(b).strip():
            return b
    return None


def target_para(d, text):
    """"the heading “X”" hoặc "the document title"."""
    m = re.search(r"heading [“\"]([^”\"]+)[”\"]", text)
    if m:
        p = find_para(d, m.group(1))
        return p if p is not None else heading_para(d, m.group(1))
    if "document title" in text:
        return title_para(d)
    return None


def documents_dirs(path: Path) -> list[Path]:
    home = Path.home()
    out = [home / "Documents"]
    out += sorted(home.glob("OneDrive*/Documents")) + sorted(home.glob("OneDrive*/Tài liệu"))
    prof = os.environ.get("USERPROFILE")
    if prof:
        out.append(Path(prof) / "Documents")
    out.append(Path(path).parent)
    seen, res = set(), []
    for p in out:
        if p not in seen:
            seen.add(p)
            res.append(p)
    return res


def session_start(path: Path) -> float | None:
    m = re.search(r"_(\d{8}_\d{6})_\d+$", Path(path).parent.name)
    if not m:
        return None
    try:
        return datetime.strptime(m.group(1), "%Y%m%d_%H%M%S").timestamp()
    except ValueError:
        return None


def saved_copy(path: Path, name: str, exts) -> Path | None:
    start = session_start(path)
    for folder in documents_dirs(path):
        for ext in exts:
            f = folder / f"{name}{ext}"
            if f.is_file() and (start is None or f.stat().st_mtime >= start - 5):
                return f
    return None


# ====================================================================== snapshot file gốc


def _sec_snap(d, text):
    name = sec_name(text)
    return section(d, name) if name else [d.body]


@snap("pic_size", "pic_change", "pic_insert")
def _(d, text):
    drs = sec_objects(d, text, "pic")
    out = {"n": len(drs)}
    if drs:
        out["cx"], out["cy"] = extent(drs[0])
        out["img"] = md5(pic_image(d, drs[0]))
    return out


@snap("chart_insert")
def _(d, text):
    return {"n": len(sec_objects(d, text, "chart"))}


@snap("chart_data")
def _(d, text):
    parts = chart_parts(d, sec_objects(d, text, "chart"))
    v = _chart_value(d, parts[0], text) if parts else None
    return {"v": v}


@snap("chart_style", "chart_colors", "chart_layout")
def _(d, text):
    parts = chart_parts(d, sec_objects(d, text, "chart"))
    return {"fp": _chart_fp(d, parts[0]) if parts else ""}


@snap("sa_colors", "sa_style", "sa_layout", "sa_effect")
def _(d, text):
    drs = sec_objects(d, text, "dgm")
    if not drs:
        return {}
    parts = dgm_parts(d, drs[0])
    return {k: _unique_id(d, parts.get(k)) for k in ("lo", "qs", "cs")}


@snap("comment_delete", "comment_add", "comment_reply", "comment_resolve")
def _(d, text):
    root = d.xml("word/commentsExtended.xml")
    done = 0 if root is None else sum(1 for c in root.iter() if _attr(c, "done", "w15") == "1")
    return {"n": len(comments(d)), "done": done}


@snap("delete_col", "split_cell", "table_insert_col")
def _(d, text):
    tbls = sec_tables(d, sec_name(text))
    return {"cols": len(tbls[0].find(W + "tr").findall(W + "tc")) if tbls else 0}


@snap("split_table", "table_to_text")
def _(d, text):
    tbls = sec_tables(d, sec_name(text))
    head = [ptext(p) for p in tbls[0].find(W + "tr").iter(P)] if tbls else []
    return {"n": len(tbls), "head": head}


@snap("info_to_text")
def _(d, text):
    q = quotes(text)
    for tb in d.body.iter(TBL):
        rows = tb.findall(W + "tr")
        if rows and q and norm(ptext(next(rows[0].iter(P)))) == norm(q[0]):
            return {"rows": [[ptext(tc) for tc in r.findall(W + "tc")] for r in rows]}
    return {}


@snap("replace_all", "wildcard_replace")
def _(d, text):
    q = quotes(text)
    body = doc_text(d)
    return {"n": sum(body.count(x) for x in q[:-1]), "n_new": body.count(q[-1])}


@snap("find_delete")
def _(d, text):
    q = quotes(text)
    return {"n": len(re.findall(rf"\b{re.escape(q[0])}\b", doc_text(d)))} if q else {}


@snap("replace_style")
def _(d, text):
    names = re.findall(r"(Heading \d)", text)
    return {"n": sum(1 for p in all_paras(d.body) if pstyle(d, p) == norm(names[0]))} if names else {}


@snap("track_accept", "track_reject")
def _(d, text):
    return revisions(d)


@snap("style_set", "para_spacing_doc", "custom_styleset")
def _(d, text):
    return {"fp": _styles_fp(d)}


@snap("index")
def _(d, text):
    q = quotes(text)
    return {"n": doc_text(d).count(q[0]) if q else 0}


@snap("sa_add")
def _(d, text):
    drs = sec_objects(d, text, "dgm")
    return {"n": len(_dgm_texts(d, drs[0])) if drs else 0}


@snap("wordart")
def _(d, text):
    p = title_para(d)
    return {"title": ptext(p).strip() if p is not None else ""}


@snap("replace_color")
def _(d, text):
    return {"n": sum(1 for c in d.body.iter(W + "color") if (_attr(c, "val") or "").upper() == "C00000")}


def snapshot(path, dang, de_bai) -> dict:
    """Ghi lại thông tin file gốc mà hàm chấm cần (số ảnh, kích thước, số comment…)."""
    fn = SNAPS.get(dang)
    if fn is None:
        return {}
    try:
        return fn(Doc(path), de_bai) or {}
    except Exception:  # file gốc lạ: bỏ qua, hàm chấm sẽ chấm lỏng hơn
        return {}


# ====================================================================== CHẤM: 1. Quản lý tài liệu


@grader("prop_category")
def _(d, t, g):
    return norm(core_prop(d, "category")) == norm(quotes(t)[0])


@grader("prop_subject")
def _(d, t, g):
    return norm(core_prop(d, "subject")) == norm(quotes(t)[0])


@grader("prop_tags")
def _(d, t, g):
    tags = [norm(x) for x in re.split(r"[;,]", core_prop(d, "keywords"))]
    return norm(quotes(t)[0]) in tags


@grader("prop_title")
def _(d, t, g):
    return norm(core_prop(d, "title")) == norm(quotes(t)[0])


@grader("prop_company")
def _(d, t, g):
    m = re.search(r"property of the document to (.+?)\.?$", t)
    return bool(m) and norm(app_prop(d, "Company")) == norm(m.group(1).strip("“”\""))


@grader("inspect_props")
def _(d, t, g):
    st = settings(d)
    if st is not None and _on(st.find(W + "removePersonalInformation")):
        return True
    return not any(core_prop(d, k).strip() for k in ("creator", "lastModifiedBy", "title", "subject",
                                                      "keywords", "category")) and not app_prop(d, "Company")


@grader("inspect_hf")
def _(d, t, g):
    for part in d.glob(r"word/(header|footer)\d*\.xml"):
        root = d.xml(part)
        referenced = any(d.rels("word/document.xml").get(ref.get("{%s}id" % NS["r"])) == part
                         for ref in d.body.iter(W + "headerReference", W + "footerReference"))
        if referenced and (ptext_all(root).strip() or drawings([root]) or root.find(".//v:shape", NS) is not None
                           or root.find(".//w:pict", NS) is not None):
            return False
    return True


def ptext_all(root) -> str:
    return "\n".join(ptext(p) for p in root.iter(P))


@grader("access_table")
def _(d, t, g):
    tbls = list(d.body.iter(TBL))
    for tb in tbls:
        tr = tb.find(W + "tr")
        if tr is None or not _on(tr.find(f"{W}trPr/{W}tblHeader")):
            return False
    return bool(tbls)


@grader("access_alt", "alt_text")
def _(d, t, g):
    val = norm(quotes(t)[-1])
    drs = sec_objects(d, t, "pic") if sec_name(t) else drawings([d.body], "pic")
    return any(norm(docpr(dr).get("descr")) == val for dr in drs if docpr(dr) is not None)


@grader("compat")
def _(d, t, g):
    st = settings(d)
    for cs in st.iter(W + "compatSetting") if st is not None else []:
        if _attr(cs, "name") == "compatibilityMode":
            return _int(_attr(cs, "val"), 0) >= 15
    return False


def _save_as(ext_list):
    def fn(d, t, g):
        name = quotes(t)[0]
        return saved_copy(d.path, name, ext_list) is not None
    return fn


GRADERS["save_txt"] = _save_as([".txt"])
GRADERS["save_pdf"] = _save_as([".pdf"])
GRADERS["save_template"] = _save_as([".dotx"])
GRADERS["save_rtf"] = _save_as([".rtf"])


@grader("save_97")
def _(d, t, g):
    f = saved_copy(d.path, quotes(t)[0], [".doc"])
    return f is not None and f.read_bytes()[:4] == b"\xd0\xcf\x11\xe0"


def _pg_mar(d):
    out = []
    for sp in sect_prs(d):
        m = sp.find(W + "pgMar")
        if m is not None:
            out.append({k: _int(_attr(m, k), 0) for k in ("top", "bottom", "left", "right")})
    return out


@grader("margins")
def _(d, t, g):
    v = inches(t)
    if len(v) < 2:
        return None
    tb, lr = v[0] * TW_IN, v[1] * TW_IN
    mars = _pg_mar(d)
    return bool(mars) and all(_near(m["top"], tb, 20) and _near(m["bottom"], tb, 20) and
                              _near(m["left"], lr, 20) and _near(m["right"], lr, 20) for m in mars)


MARGIN_PRESETS = {"narrow": (0.5, 0.5), "moderate": (1, 0.75), "wide": (1, 2), "normal": (1, 1),
                  "mirrored": (1, None)}


@grader("margins_preset")
def _(d, t, g):
    m = re.search(r"Apply the (\w+) margin", t, re.I)
    key = m.group(1).casefold() if m else ""
    if key not in MARGIN_PRESETS:
        return None
    tb, lr = MARGIN_PRESETS[key]
    mars = _pg_mar(d)
    if not mars:
        return False
    if key == "mirrored":
        st = settings(d)
        return st is not None and _on(st.find(W + "mirrorMargins")) and all(
            _near(x["top"], 1440, 20) and _near(x["left"], 1.25 * TW_IN, 20) for x in mars)
    return all(_near(x["top"], tb * TW_IN, 20) and _near(x["bottom"], tb * TW_IN, 20) and
               _near(x["left"], lr * TW_IN, 20) and _near(x["right"], lr * TW_IN, 20) for x in mars)


def _landscape(sp) -> bool:
    sz = sp.find(W + "pgSz")
    if sz is None:
        return False
    return _int(_attr(sz, "w"), 0) > _int(_attr(sz, "h"), 0)


@grader("orient_all")
def _(d, t, g):
    sps = sect_prs(d)
    return bool(sps) and all(_landscape(sp) for sp in sps)


PAPER = {"a4": (11906, 16838), "a5": (8391, 11906), "letter": (12240, 15840), "legal": (12240, 20160),
         "executive": (10440, 15120), "a3": (16838, 23811)}


@grader("paper")
def _(d, t, g):
    m = re.search(r"paper size of the document to ([\w ]+)", t)
    key = m.group(1).strip().casefold() if m else ""
    if key not in PAPER:
        return None
    want = sorted(PAPER[key])
    sps = sect_prs(d)
    for sp in sps:
        sz = sp.find(W + "pgSz")
        got = sorted([_int(_attr(sz, "w"), 0), _int(_attr(sz, "h"), 0)]) if sz is not None else [0, 0]
        if not (_near(got[0], want[0], 60) and _near(got[1], want[1], 60)):
            return False
    return bool(sps)


def _theme(d):
    for n in d.glob(r"word/theme/theme\d*\.xml"):
        return d.xml(n)
    return None


@grader("doc_theme")
def _(d, t, g):
    m = re.search(r"Apply the (.+?) theme", t, re.I)
    th = _theme(d)
    return bool(m) and th is not None and norm(th.get("name")).startswith(norm(m.group(1)))


@grader("theme_colors")
def _(d, t, g):
    m = re.search(r"theme colors of the document to (.+?) and the theme fonts to (.+?)\.?$", t)
    th = _theme(d)
    if not m or th is None:
        return False
    cs = th.find(".//a:clrScheme", NS)
    fs = th.find(".//a:fontScheme", NS)
    ok_c = cs is not None and norm(cs.get("name")) == norm(m.group(1))
    major = fs.find("a:majorFont/a:latin", NS) if fs is not None else None
    ok_f = fs is not None and (norm(fs.get("name")) == norm(m.group(2)) or
                               (major is not None and norm(major.get("typeface")) == norm(m.group(2))))
    return ok_c and ok_f


@grader("custom_colors")
def _(d, t, g):
    m = re.search(r"uses (.+?) \(Standard Colors\) as the Accent (\d)", t)
    th = _theme(d)
    if not m or th is None:
        return False
    cs = th.find(".//a:clrScheme", NS)
    if cs is None or norm(cs.get("name")) != norm(quotes(t)[0]):
        return False
    acc = cs.find(f"a:accent{m.group(2)}", NS)
    return dml_color_ok(acc, color_spec(m.group(1)))


@grader("custom_fonts")
def _(d, t, g):
    m = re.search(r"uses (.+?) as the heading font and (.+?) as the body font", t)
    th = _theme(d)
    if not m or th is None:
        return False
    fs = th.find(".//a:fontScheme", NS)
    if fs is None:
        return False
    major, minor = fs.find("a:majorFont/a:latin", NS), fs.find("a:minorFont/a:latin", NS)
    return (norm(fs.get("name")) == norm(quotes(t)[0]) and major is not None and minor is not None and
            norm(major.get("typeface")) == norm(m.group(1)) and norm(minor.get("typeface")) == norm(m.group(2)))


def _styles_fp(d) -> str:
    root = d.xml("word/styles.xml")
    if root is None:
        return ""
    keep = [root.find("w:docDefaults", NS)]
    for sid in ("Normal", "Title", "Heading1", "Heading2"):
        keep.append(d.styles().get(sid))
    return md5(b"".join(etree.tostring(x) for x in keep if x is not None))


@grader("style_set")
def _(d, t, g):
    return bool(g.get("fp")) and _styles_fp(d) != g["fp"]


SPACING_PRESETS = {"double": 480, "relaxed": 360}


@grader("para_spacing_doc")
def _(d, t, g):
    m = re.search(r"entire document to (\w+)", t)
    key = m.group(1).casefold() if m else ""
    root = d.xml("word/styles.xml")
    if root is None:
        return False
    if key in SPACING_PRESETS:
        ps = [p for p in blocks(d) if p.tag == P and ptext(p).strip() and pstyle(d, p) == "normal"]
        return bool(ps) and all(_near(ppr_attr(d, p, "w:spacing", "line"), SPACING_PRESETS[key], 12)
                                for p in ps)
    return bool(g.get("fp")) and _styles_fp(d) != g["fp"]


def _header_parts(d, kind):
    rels = d.rels("word/document.xml")
    out = []
    for ref in d.body.iter(W + f"{kind}Reference"):
        part = rels.get(ref.get("{%s}id" % NS["r"]))
        if part:
            out.append((_attr(ref, "type") or "default", part))
    return out


def _has_content(d, part):
    root = d.xml(part)
    return root is not None and (ptext_all(root).strip() or drawings([root]) or
                                 root.find(".//w:sdt", NS) is not None or root.find(".//w:pict", NS) is not None)


def _hf_prints(d, kind) -> list[str]:
    return [md5(d.files.get(p, b"")) for _typ, p in _header_parts(d, kind)]


@snap("header_gallery", "footer_gallery")
def _(d, text):
    return {"old": _hf_prints(d, "header") + _hf_prints(d, "footer")}


@grader("header_gallery", "footer_gallery")
def _(d, t, g):
    kind = "header" if " header" in t else "footer"
    parts = _header_parts(d, kind)
    old = set(g.get("old", []))
    if not any(typ == "default" and _has_content(d, p) and md5(d.files.get(p, b"")) not in old
               for typ, p in parts):
        return False
    title_pg = any(_on(sp.find(W + "titlePg")) for sp in sect_prs(d))
    if "except page 1" in t:
        return title_pg
    return not title_pg or any(typ == "first" and _has_content(d, p) for typ, p in parts)


@grader("page_numbers")
def _(d, t, g):
    ok = False
    for kind in ("header", "footer"):
        for _typ, part in _header_parts(d, kind):
            root = d.xml(part)
            if root is not None and any(re.match(r"PAGE\b", f) for f in fields(root)):
                ok = True
    m = re.search(r"start at (\d+)", t)
    if ok and m:
        return any(_attr(sp.find(W + "pgNumType"), "start") == m.group(1) for sp in sect_prs(d))
    return ok


@grader("watermark")
def _(d, t, g):
    q = quotes(t)
    if q:
        want = q[0]
    else:
        m = re.search(r"Add the (.+?) watermark", t)
        want = re.sub(r"\s+\d+$", "", m.group(1)) if m else ""
    for kind in ("header",):
        for _typ, part in _header_parts(d, kind):
            raw = d.raw(part)
            if re.search(r'string="' + re.escape(want) + '"', raw, re.I) or (
                    "PowerPlusWaterMark" in raw and norm(want) in norm(re.sub(r"<[^>]+>", " ", raw))):
                return True
    return False


@grader("page_color")
def _(d, t, g):
    m = re.search(r"page color of the document to (.+?)\.?$", t)
    bg = d.doc.find(W + "background")
    return bool(m) and color_ok(bg, color_spec(m.group(1)))


BORDER_SZ = {"½": 4, "¾": 6, "1": 8, "1 ½": 12, "2 ¼": 18, "3": 24, "4 ½": 36, "6": 48}


def _border_sz(text):
    m = re.search(r"(\d(?: [½¼¾])?|[½¾]) pt", text)
    return BORDER_SZ.get(m.group(1)) if m else None


@grader("page_border")
def _(d, t, g):
    m = re.search(r"Add a (.+?) pt (.+?) Box page border", t)
    if not m:
        return None
    sz, spec = _border_sz(t), color_spec(m.group(2))
    for sp in sect_prs(d):
        pb = sp.find(W + "pgBorders")
        if pb is None:
            continue
        sides = [pb.find(W + s) for s in ("top", "left", "bottom", "right")]
        if any(s is None or _attr(s, "val") in (None, "nil", "none") for s in sides):
            continue
        if sz and not all(_near(_attr(s, "sz"), sz, 1) for s in sides):
            continue
        if not color_ok(sides[0], spec, "border"):
            continue
        first = _attr(pb, "display") == "firstPage"
        if ("first page only" in t) == first:
            return True
    return False


@grader("hyphenation")
def _(d, t, g):
    st = settings(d)
    return st is not None and _on(st.find(W + "autoHyphenation"))


@grader("hyperlink_url")
def _(d, t, g):
    phrase = quotes(t)[-1]
    m = re.search(r"links to (https?://\S+?)\.?(?:\s|$)", t)
    url = m.group(1).rstrip("/.") if m else ""
    rels = d.rels("word/document.xml")
    for h in d.body.iter(W + "hyperlink"):
        target = rels.get(h.get("{%s}id" % NS["r"]), "")
        if norm(phrase) in norm(ptext(h)) and target.rstrip("/").casefold() == url.casefold():
            if "ScreenTip" in t:
                tip = quotes(t.split("ScreenTip")[1])
                return bool(tip) and norm(_attr(h, "tooltip")) == norm(tip[0])
            return True
    for f_para in all_paras(d.body):
        if any(url.casefold() in f.casefold() and f.upper().startswith("HYPERLINK") for f in fields(f_para)) \
                and norm(phrase) in norm(ptext(f_para)):
            return True
    return False


@grader("hyperlink_place")
def _(d, t, g):
    q = quotes(t)
    phrase, target = q[1], q[2]
    h_para = heading_para(d, target)
    names = {_attr(b, "name") for b in h_para.iter(W + "bookmarkStart")} if h_para is not None else set()
    if h_para is not None:
        prev = h_para.getprevious()
        while prev is not None and prev.tag == W + "bookmarkStart":
            names.add(_attr(prev, "name"))
            prev = prev.getprevious()
    for h in d.body.iter(W + "hyperlink"):
        if norm(phrase) in norm(ptext(h)) and _attr(h, "anchor"):
            return _attr(h, "anchor") in names or not names
    return False


@grader("bookmark")
def _(d, t, g):
    q = quotes(t)
    start, name = q[0], q[-1]
    for bm in d.body.iter(W + "bookmarkStart"):
        if norm(_attr(bm, "name")) != norm(name):
            continue
        p = bm if bm.tag == P else next((a for a in bm.iterancestors(P)), None)
        if p is None:
            p = bm.getnext()
            while p is not None and p.tag != P:
                p = p.getnext()
        if p is not None and norm(ptext(p)).startswith(norm(start.rstrip(" …"))):
            return True
    return False


@grader("encrypt")
def _(d, t, g):
    return False     # file đọc được như ZIP → chưa mã hóa (file mã hóa xử lý trong grade())


@grader("mark_final")
def _(d, t, g):
    raw = d.raw("docProps/custom.xml")
    if re.search(r'name="_MarkAsFinal"[^>]*>\s*<vt:bool>(true|1)</vt:bool>', raw, re.I):
        return True
    return norm(core_prop(d, "contentStatus")) == "final"


def _protection(d):
    st = settings(d)
    return st.find(W + "documentProtection") if st is not None else None


@grader("restrict_comments")
def _(d, t, g):
    pr = _protection(d)
    return pr is not None and _attr(pr, "edit") == "comments" and _on(pr, "enforcement")


@grader("restrict_format")
def _(d, t, g):
    pr = _protection(d)
    return pr is not None and _on(pr, "formatting") and not (
        "Do not enforce" in t and _attr(pr, "enforcement") and _on(pr, "enforcement"))


def _track_on(d):
    st = settings(d)
    return st is not None and _on(st.find(W + "trackRevisions"))


@grader("lock_tracking", "track_lock")
def _(d, t, g):
    pr = _protection(d)
    ok = (pr is not None and _attr(pr, "edit") == "trackedChanges" and _on(pr, "enforcement") and
          bool(_attr(pr, "hashValue") or _attr(pr, "hash") or _attr(pr, "cryptProviderType") or
               _attr(pr, "algIdExt") or _attr(pr, "algorithmName")))
    return ok and (_track_on(d) if "Turn on Track Changes" in t else True)


# ====================================================================== 2. Chữ, đoạn, section


@grader("find_delete")
def _(d, t, g):
    word = quotes(t)[0]
    return not re.search(rf"\b{re.escape(word)}\b", doc_text(d))


@grader("replace_all")
def _(d, t, g):
    old, new = quotes(t)[0], quotes(t)[1]
    body = doc_text(d)
    if old in body:
        return False
    if g.get("n") is not None:
        return body.count(new) >= g.get("n_new", 0) + g["n"]
    return new in body


@grader("wildcard_replace")
def _(d, t, g):
    q = quotes(t)
    body = doc_text(d)
    if any(x in body for x in q[:-1]):
        return False
    return body.count(q[-1]) >= (g.get("n_new", 0) if g else 1) and q[-1] in body


@grader("symbol")
def _(d, t, g):
    m = re.search(r"Use the (.+?) font and character code [“\"](\w+)[”\"]", t)
    if not m:
        return None
    font, code = m.group(1), int(m.group(2), 16 if "Emoji" in m.group(1) else 10)
    if "document title" in t:
        p = title_para(d)
        paras = [p] if p is not None else []
    else:
        phrase = quotes(t)[1]
        paras = [p for p in all_paras(d.body) if norm(phrase) in norm(ptext(p))]
    chars = {f"{code:04X}", f"{0xF000 + code:04X}"}
    for p in paras:
        for s in p.iter(W + "sym"):
            if norm(_attr(s, "font")) == norm(font) and (_attr(s, "char") or "").upper() in chars:
                return True
        for r in text_runs(p):
            f = r.find(f"{W}rPr/{W}rFonts")
            txt = "".join(x.text or "" for x in r if x.tag == T)
            if f is not None and norm(_attr(f, "ascii") or _attr(f, "hAnsi")) == norm(font) and any(
                    f"{ord(c):04X}" in chars for c in txt):
                return True
    return False


@grader("special_char")
def _(d, t, g):
    m = re.search(r"\((.)\)", t)
    org = quotes(t)[-1]
    body = doc_text(d)
    i = body.find(org)
    return bool(m) and i >= 0 and body[i + len(org):i + len(org) + 1] == m.group(1)


@grader("format_painter")
def _(d, t, g):
    ps = sec_paras(d, sec_name(t))
    if len(ps) < 2:
        return False

    def fp(p):
        r = text_runs(p)[0] if text_runs(p) else None
        if r is None:
            return None
        out = []
        for name in ("i", "b"):
            out.append(_on(rpr_find(d, r, p, name)))
        c = rpr_find(d, r, p, "color")
        out.append((_attr(c, "val") or "").upper())
        out.append(_attr(rpr_find(d, r, p, "sz"), "val"))
        f = rpr_find(d, r, p, "rFonts")
        out.append(norm(_attr(f, "ascii") or _attr(f, "hAnsi")))
        return out
    a, b = fp(ps[0]), fp(ps[1])
    return a is not None and a == b and (a[0] or a[2] or a[4])


@grader("line_spacing_doc")
def _(d, t, g):
    m = re.search(r"line spacing to ([\d.]+) lines", t)
    want = float(m.group(1)) * 240
    ps = [p for p in all_paras(d.body) if ptext(p).strip()]
    good = [p for p in ps if _near(ppr_attr(d, p, "w:spacing", "line"), want, 3) and
            (ppr_attr(d, p, "w:spacing", "lineRule") or "auto") == "auto"]
    return bool(ps) and len(good) >= 0.9 * len(ps)


@grader("line_spacing_exact")
def _(d, t, g):
    v = float(re.search(r"exactly (\d+(?:\.\d+)?) pt", t).group(1)) * 20
    ps = sec_paras(d, sec_name(t))[:2]
    return len(ps) == 2 and all(_near(ppr_attr(d, p, "w:spacing", "line"), v, 2) and
                                ppr_attr(d, p, "w:spacing", "lineRule") == "exact" for p in ps)


@grader("para_spacing")
def _(d, t, g):
    m = re.search(r"before .*? to (\d+) pt and the spacing after to (\d+) pt", t)
    bf, af = int(m.group(1)) * 20, int(m.group(2)) * 20
    ps = sec_paras(d, sec_name(t))[:2]
    return len(ps) == 2 and all(_near(ppr_attr(d, p, "w:spacing", "before") or 0, bf, 2) and
                                _near(ppr_attr(d, p, "w:spacing", "after") or 0, af, 2) for p in ps)


@grader("indent")
def _(d, t, g):
    v = inches(t)[0] * TW_IN
    ps = sec_paras(d, sec_name(t))[:2]
    if len(ps) < 2:
        return False
    if "first line" in t:
        return all(_near(ppr_attr(d, p, "w:ind", "firstLine"), v, 15) for p in ps)
    return all(_near(ppr_attr(d, p, "w:ind", "left") or ppr_attr(d, p, "w:ind", "start"), v, 15) for p in ps)


@grader("clear_format")
def _(d, t, g):
    p = find_para(d, quotes(t)[0], start=True)
    if p is None:
        return False
    if pstyle(d, p) not in ("normal",) or _attr(p.find(f"{W}pPr/{W}jc"), "val") in ("center", "right", "both"):
        return False
    for r in text_runs(p):
        rpr = r.find(W + "rPr")
        if rpr is not None and any(rpr.find(W + n) is not None and (n != "b" and n != "i" or _on(rpr.find(W + n)))
                                   for n in ("b", "i", "color", "sz", "rFonts", "highlight", "u")):
            return False
    return True


@grader("modify_style")
def _(d, t, g):
    m = re.search(r"Modify the (.+?) style so that it uses (\d+) pt, (bold, )?(italic, )?(.+?) font color", t)
    if not m:
        return None
    st = d.style_by_name(m.group(1))
    rpr = st.find(W + "rPr") if st is not None else None
    if rpr is None:
        return False
    return (_attr(rpr.find(W + "sz"), "val") == str(int(m.group(2)) * 2) and
            (not m.group(3) or _on(rpr.find(W + "b"))) and (not m.group(4) or _on(rpr.find(W + "i"))) and
            color_ok(rpr.find(W + "color"), color_spec(m.group(5))))


@grader("replace_style")
def _(d, t, g):
    names = re.findall(r"(Heading \d)", t)
    if len(names) < 2:
        return None
    styles = [pstyle(d, p) for p in all_paras(d.body) if ptext(p).strip()]
    return norm(names[0]) not in styles and styles.count(norm(names[1])) >= max(1, g.get("n", 1))


@grader("para_style")
def _(d, t, g):
    m = re.search(r"Apply the (.+?) style to the paragraph", t, re.I)
    p = find_para(d, quotes(t)[0], start=True)
    return p is not None and pstyle(d, p) == norm(m.group(1))


@grader("heading_style")
def _(d, t, g):
    m = re.search(r"Apply the (.+?) style to the heading", t, re.I)
    p = find_para(d, quotes(t)[0])
    return p is not None and pstyle(d, p) == norm(m.group(1))


@grader("char_style")
def _(d, t, g):
    m = re.search(r"apply the (.+?) style to it", t, re.I)
    return _runs_have_style(d, quotes(t)[0], m.group(1))


def _runs_have_style(d, sentence, style) -> bool:
    for p in all_paras(d.body):
        if norm(sentence.rstrip(".")) in norm(ptext(p)):
            runs = runs_covering(p, sentence.rstrip("."))
            if runs and all(d.style_name(_attr(r.find(f"{W}rPr/{W}rStyle"), "val")) == norm(style)
                            for r in runs):
                return True
            if pstyle(d, p) == norm(style):
                return True
    return False


@grader("char_style_new")
def _(d, t, g):
    q = quotes(t)
    name, sentence = q[0], q[1]
    m = re.search(r"uses bold, (.+?) font color", t)
    st = d.style_by_name(name)
    if st is None or st.get(W + "type") not in ("character", "paragraph"):
        return False
    rpr = st.find(W + "rPr")
    if rpr is None or not _on(rpr.find(W + "b")):
        return False
    if m and not color_ok(rpr.find(W + "color"), color_spec(m.group(1))):
        return False
    return _runs_have_style(d, sentence, name)


@grader("para_style_new")
def _(d, t, g):
    q = quotes(t)
    name, start = q[0], q[1]
    m = re.search(r"uses the (.+?) font, (\d+) pt, italic, with (\d+) pt of spacing after", t)
    st = d.style_by_name(name)
    if st is None or st.get(W + "type") != "paragraph" or not m:
        return False
    rpr, ppr = st.find(W + "rPr"), st.find(W + "pPr")
    f = rpr.find(W + "rFonts") if rpr is not None else None
    ok = (f is not None and norm(_attr(f, "ascii") or _attr(f, "hAnsi")) == norm(m.group(1)) and
          _attr(rpr.find(W + "sz"), "val") == str(int(m.group(2)) * 2) and _on(rpr.find(W + "i")) and
          ppr is not None and _near(_attr(ppr.find(W + "spacing"), "after"), int(m.group(3)) * 20, 2))
    p = find_para(d, start, start=True)
    return ok and p is not None and pstyle(d, p) == norm(name)


@grader("change_case")
def _(d, t, g):
    head = quotes(t)[0]
    mode = t.rsplit(" to ", 1)[-1].strip().rstrip(".")
    words = head.split()
    want = {"UPPERCASE": head.upper(), "lowercase": head.lower(),
            "Capitalize Each Word": " ".join(w[:1].upper() + w[1:].lower() for w in words),
            "Sentence case": head[:1].upper() + head[1:].lower()}.get(mode)
    if want is None:
        return None
    for p in blocks(d):
        if p.tag == P and norm(ptext(p)) == norm(head):
            txt = " ".join(ptext(p).split())
            if txt == want:
                return True
            if mode == "UPPERCASE" and all(_on(rpr_find(d, r, p, "caps")) for r in text_runs(p)):
                return True
    return False


@grader("font_format")
def _(d, t, g):
    m = re.search(r"to (.+?), (\d+) pt", t)
    p = find_para(d, quotes(t)[0])
    if p is None or not m:
        return False
    for r in text_runs(p):
        f = rpr_find(d, r, p, "rFonts")
        if norm(_attr(f, "ascii") or _attr(f, "hAnsi")) != norm(m.group(1)):
            return False
        if _attr(rpr_find(d, r, p, "sz"), "val") != str(int(m.group(2)) * 2):
            return False
    return bool(text_runs(p))


@grader("highlight")
def _(d, t, g):
    sentence = quotes(t)[-1]
    m = re.search(r"in (Bright Green|Yellow|Turquoise|Pink|Blue|Red|Green|Gray.*?)\.?$", t)
    want = {"bright green": "green", "yellow": "yellow", "turquoise": "cyan", "pink": "magenta"}.get(
        norm(m.group(1)) if m else "", "green")
    for p in all_paras(d.body):
        if norm(sentence.rstrip(".")) in norm(ptext(p)):
            runs = runs_covering(p, sentence.rstrip("."))
            return bool(runs) and all(_attr(r.find(f"{W}rPr/{W}highlight"), "val") == want or
                                      _attr(r.find(f"{W}rPr/{W}shd"), "fill") in ("00FF00",) for r in runs)
    return False


@grader("para_shading")
def _(d, t, g):
    m = re.search(r"Apply (.+?) shading to the paragraph", t)
    p = find_para(d, quotes(t)[0], start=True)
    if p is None or not m:
        return False
    spec = color_spec(m.group(1))
    if color_ok(p.find(f"{W}pPr/{W}shd"), spec, "fill"):
        return True
    runs = text_runs(p)
    return bool(runs) and all(color_ok(r.find(f"{W}rPr/{W}shd"), spec, "fill") for r in runs)


@grader("text_effect")
def _(d, t, g):
    m = re.search(r"Apply the (.+?) text effect", t, re.I)
    p = target_para(d, t)
    if p is None or not m:
        return False
    return _w14_effects(d, p, m.group(1))


def _w14_effects(d, p, desc) -> bool:
    need = []
    low = desc.casefold()
    if "shadow" in low:
        need.append("shadow")
    if "outline" in low:
        need.append("textOutline")
    if "glow" in low:
        need.append("glow")
    if "reflection" in low:
        need.append("reflection")
    if "gradient" in low or "fill" in low:
        need.append("textFill")
    runs = text_runs(p)
    if not runs:
        return False
    for r in runs:
        rpr = r.find(W + "rPr")
        tags = {etree.QName(x).localname for x in rpr} if rpr is not None else set()
        if not all(n in tags for n in need if n != "textFill") or not (tags & {"shadow", "textOutline", "glow",
                                                                               "reflection", "textFill"}):
            return False
    return True


@grader("keep_next")
def _(d, t, g):
    p = heading_para(d, quotes(t)[0])
    return p is not None and _on(ppr_find(d, p, "w:keepNext")) and _on(ppr_find(d, p, "w:keepLines"))


@grader("widow")
def _(d, t, g):
    ps = sec_paras(d, sec_name(t))[:2]
    return len(ps) == 2 and all(not _on(p.find(f"{W}pPr/{W}widowControl")) and
                                p.find(f"{W}pPr/{W}widowControl") is not None and
                                _on(ppr_find(d, p, "w:keepLines")) for p in ps)


@grader("page_break")
def _(d, t, g):
    h = heading_para(d, quotes(t)[0])
    if h is None:
        return False
    if _on(ppr_find(d, h, "w:pageBreakBefore")):
        return True
    for br in h.iter(W + "br"):
        if _attr(br, "type") == "page":
            return True
    prev = h.getprevious()
    while prev is not None and prev.tag == P and not ptext(prev).strip():
        if any(_attr(br, "type") == "page" for br in prev.iter(W + "br")):
            return True
        prev = prev.getprevious()
    return prev is not None and prev.tag == P and any(_attr(br, "type") == "page" for br in prev.iter(W + "br"))


@grader("sbreak_cont")
def _(d, t, g):
    sentence = quotes(t)[0]
    for p in all_paras(d.body):
        if norm(sentence.rstrip(".")) in norm(ptext(p)):
            sp = p.find(f"{W}pPr/{W}sectPr")
            if sp is not None and _attr(sp.find(W + "type"), "val") == "continuous":
                return True
            nxt = p.getnext()
            if nxt is not None and nxt.tag == P and not ptext(nxt).strip():
                sp = nxt.find(f"{W}pPr/{W}sectPr")
                if sp is not None and _attr(sp.find(W + "type"), "val") == "continuous":
                    return True
    return False


@grader("sbreak_next")
def _(d, t, g):
    h = heading_para(d, quotes(t)[0])
    if h is None:
        return False
    prev = h.getprevious()
    while prev is not None and prev.tag != P:
        prev = prev.getprevious()
    sp = prev.find(f"{W}pPr/{W}sectPr") if prev is not None else None
    return sp is not None and (_attr(sp.find(W + "type"), "val") or "nextPage") == "nextPage"


def _governing_sect(d, p):
    """sectPr quy định trang chứa đoạn p (sectPr kế tiếp sau đoạn)."""
    seen = False
    for b in d.body.iter(P, W + "sectPr"):
        if b is p:
            seen = True
        if seen and b.tag == W + "sectPr":
            return b
    return final_sect(d)


@grader("orient_sec")
def _(d, t, g):
    h = heading_para(d, quotes(t)[0])
    if h is None:
        return False
    sp = _governing_sect(d, h)
    others = [x for x in sect_prs(d) if x is not sp]
    return sp is not None and _landscape(sp) and bool(others) and any(not _landscape(x) for x in others)


@grader("columns")
def _(d, t, g):
    m = re.search(r"into (\d) columns", t)
    n = m.group(1)
    sp_v = inches(t)
    for sp in sect_prs(d):
        cols = sp.find(W + "cols")
        if cols is None or _attr(cols, "num") != n:
            continue
        if "line between" in t and not _on(cols, "sep"):
            continue
        if sp_v and not _near(_attr(cols, "space"), sp_v[0] * TW_IN, 15):
            continue
        return True
    return False


@grader("dropcap")
def _(d, t, g):
    m = re.search(r"drop (\d) lines", t)
    kind = "margin" if "In margin" in t else "drop"
    for b in section(d, sec_name(t)):
        for fp in b.iter(W + "framePr"):
            if _attr(fp, "dropCap") == kind and (not m or _attr(fp, "lines") == m.group(1)):
                return True
    return False


@grader("language")
def _(d, t, g):
    m = re.search(r"\((.+?)\)", t)
    code = {"united kingdom": "en-GB", "united states": "en-US", "australia": "en-AU", "canada": "en-CA"}.get(
        norm(m.group(1)) if m else "", "en-GB")
    root = d.xml("word/styles.xml")
    default = _attr(root.find("w:docDefaults/w:rPrDefault/w:rPr/w:lang", NS), "val") if root is not None else None
    runs = [(r, p) for p in all_paras(d.body) for r in text_runs(p)]
    good = 0
    for r, p in runs:
        lang = _attr(rpr_find(d, r, p, "lang"), "val") or default
        good += lang == code
    return bool(runs) and good >= 0.9 * len(runs)


@grader("line_numbers")
def _(d, t, g):
    for sp in sect_prs(d):
        ln = sp.find(W + "lnNumType")
        if ln is None:
            return False
        restart = _attr(ln, "restart") or "newPage"
        want = "newPage" if "each page" in t else "newSection" if "each section" in t else "continuous"
        if restart != want:
            return False
    return bool(sect_prs(d))


@grader("replace_color")
def _(d, t, g):
    m = re.search(r"formatted with the (.+?) font color and change the font color to (.+?)\.?$", t)
    if not m:
        return None
    old, new = color_spec(m.group(1)), color_spec(m.group(2))
    runs = [r for p in all_paras(d.body) for r in text_runs(p)]
    if any(color_ok(r.find(f"{W}rPr/{W}color"), old) for r in runs):
        return False
    return sum(1 for r in runs if color_ok(r.find(f"{W}rPr/{W}color"), new)) >= max(1, g.get("n", 1))


# ====================================================================== 3. Bảng và danh sách


def _table(d, t):
    name = sec_name(t)
    if name:
        tbls = sec_tables(d, name)
        return tbls[0] if tbls else None
    q = quotes(t)
    for tb in d.body.iter(TBL):
        first = next(tb.iter(P), None)
        if q and first is not None and norm(ptext(first)) == norm(q[0]):
            return tb
    return None


def _rows(tb):
    return tb.findall(W + "tr") if tb is not None else []


def _cells(tr):
    return [c for c in tr.iter(W + "tc") if c.getparent().tag in (W + "tr", W + "sdtContent")]


def _ctext(tc):
    return " ".join(ptext(p) for p in tc.iter(P)).strip()


@grader("add_row")
def _(d, t, g):
    vals = quotes(t)[1:]
    rows = _rows(_table(d, t))
    return bool(rows) and [norm(_ctext(c)) for c in _cells(rows[-1])] == [norm(v) for v in vals]


@grader("table_insert_col")
def _(d, t, g):
    q = quotes(t)
    after, new = q[1], q[2]
    rows = _rows(_table(d, t))
    if not rows:
        return False
    head = [norm(_ctext(c)) for c in _cells(rows[0])]
    return norm(after) in head and head.index(norm(after)) + 1 < len(head) and \
        head[head.index(norm(after)) + 1] == norm(new) and len(head) == g.get("cols", len(head) - 1) + 1


@grader("delete_col")
def _(d, t, g):
    col = quotes(t)[1]
    rows = _rows(_table(d, t))
    if not rows:
        return False
    head = [norm(_ctext(c)) for c in _cells(rows[0])]
    return norm(col) not in head and (not g.get("cols") or len(head) == g["cols"] - 1)


@grader("merge_cells")
def _(d, t, g):
    tb = _table(d, t)
    rows = _rows(tb)
    if not rows:
        return False
    ncols = len(tb.findall(f"{W}tblGrid/{W}gridCol"))
    cells = _cells(rows[0])
    return len(cells) == 1 and (ncols <= 1 or _int(_attr(cells[0].find(f"{W}tcPr/{W}gridSpan"), "val"), 1) == ncols)


@grader("split_cell")
def _(d, t, g):
    rows = _rows(_table(d, t))
    head = _cells(rows[0]) if rows else []
    return bool(head) and len(head) == g.get("cols", 0) + 1 and any(
        norm(_ctext(c)) == norm(quotes(t)[1]) for c in head)


@grader("split_table")
def _(d, t, g):
    start = quotes(t)[1]
    tbls = sec_tables(d, sec_name(t))
    if len(tbls) < 2:
        return False
    return any(_rows(tb) and norm(_ctext(_cells(_rows(tb)[0])[0])).startswith(norm(start)) for tb in tbls[1:])


@grader("table_to_text", "info_to_text")
def _(d, t, g):
    sep = "," if "commas" in t else "\t"
    if sec_name(t):
        if sec_tables(d, sec_name(t)):
            return False
        head = g.get("head")
        rows = [ptext(p) for p in sec_paras(d, sec_name(t)) if sep in ptext(p)]
        if head:
            return any(norm(x.split(sep)[0]) == norm(head[0]) for x in rows)
        return bool(rows)
    q = quotes(t)
    if _table(d, t) is not None:
        return False
    for p in all_paras(d.body):
        x = ptext(p)
        if norm(x).startswith(norm(q[0])) and sep in x:
            return True
    return False


@grader("text_to_table")
def _(d, t, g):
    first = quotes(t)[1]
    m = re.search(r"to a (\w+)-column table", t)
    n = {"two": 2, "three": 3, "four": 4}.get(m.group(1) if m else "two", 2)
    for tb in sec_tables(d, sec_name(t)):
        rows = _rows(tb)
        if rows and norm(_ctext(_cells(rows[0])[0])) == norm(first) and len(_cells(rows[0])) == n:
            return True
    return False


@grader("insert_table")
def _(d, t, g):
    m = re.search(r"has (\d+) columns and (\d+) rows", t)
    head = quotes(t)[1:]
    h = heading_para(d, quotes(t)[0])
    if h is None or not m:
        return False
    nxt = h.getnext()
    while nxt is not None and nxt.tag == P and not ptext(nxt).strip():
        nxt = nxt.getnext()
    if nxt is None or nxt.tag != TBL:
        return False
    rows = _rows(nxt)
    ok = (len(rows) == int(m.group(2)) and all(len(_cells(r)) == int(m.group(1)) for r in rows) and
          [norm(_ctext(c)) for c in _cells(rows[0])] == [norm(x) for x in head])
    if ok and "Fit the table to its contents" in t:
        return _autofit_contents(nxt)
    return ok


def _autofit_contents(tb):
    lay = tb.find(f"{W}tblPr/{W}tblLayout")
    if lay is not None and _attr(lay, "type") == "fixed":
        return False
    tw = tb.find(f"{W}tblPr/{W}tblW")
    return tw is None or _attr(tw, "type") in ("auto", None) or _attr(tw, "w") == "0"


def _grid(tb):
    return [_int(_attr(c, "w")) for c in tb.findall(f"{W}tblGrid/{W}gridCol")]


@snap("autofit")
def _(d, text):
    tb = _table(d, text)
    return {"grid": _grid(tb)} if tb is not None else {}


@grader("autofit")
def _(d, t, g):
    tb = _table(d, t)
    if tb is None:
        return False
    tw = tb.find(f"{W}tblPr/{W}tblW")
    if "Window" in t:
        return tw is not None and _attr(tw, "type") == "pct" and (_attr(tw, "w") or "") in ("5000", "100%")
    if not _autofit_contents(tb):
        return False
    widths = [_attr(c.find(f"{W}tcPr/{W}tcW"), "type") for r in _rows(tb) for c in _cells(r)]
    # Word có thể giữ tcW dạng dxa nhưng tính lại độ rộng cột theo nội dung
    return all(w in (None, "auto") for w in widths) or (bool(g.get("grid")) and _grid(tb) != g["grid"])


@grader("col_width")
def _(d, t, g):
    v = inches(t)[0] * TW_IN
    tb = _table(d, t)
    if tb is None:
        return False
    grid = [_int(_attr(c, "w")) for c in tb.findall(f"{W}tblGrid/{W}gridCol")]
    cells = [_int(_attr(c.find(f"{W}tcPr/{W}tcW"), "w")) for r in _rows(tb) for c in _cells(r)]
    return bool(grid) and all(_near(x, v, 20) for x in grid) and all(x is None or _near(x, v, 20) for x in cells)


@grader("table_rowheight")
def _(d, t, g):
    v = inches(t)[0] * TW_IN
    rows = _rows(_table(d, t))
    if not rows:
        return False
    for r in rows:
        if not _near(_attr(r.find(f"{W}trPr/{W}trHeight"), "val"), v, 15):
            return False
        if any(_attr(c.find(f"{W}tcPr/{W}vAlign"), "val") != "center" for c in _cells(r)):
            return False
    return True


@grader("cell_align")
def _(d, t, g):
    m = re.search(r"to Align (Top|Center|Bottom) (Left|Center|Right)", t)
    rows = _rows(_table(d, t))
    if not rows or not m:
        return False
    v = {"Top": "top", "Center": "center", "Bottom": "bottom"}[m.group(1)]
    h = {"Left": ("left", "start", None), "Center": ("center",), "Right": ("right", "end")}[m.group(2)]
    for c in _cells(rows[0]):
        if (_attr(c.find(f"{W}tcPr/{W}vAlign"), "val") or "top") != v:
            return False
        for p in c.iter(P):
            if _attr(p.find(f"{W}pPr/{W}jc"), "val") not in h:
                return False
    return True


@grader("cell_spacing")
def _(d, t, g):
    v = inches(t)[0] * TW_IN
    tb = _table(d, t)
    cs = tb.find(f"{W}tblPr/{W}tblCellSpacing") if tb is not None else None
    w = _int(_attr(cs, "w"))
    return w is not None and (_near(w, v, 4) or _near(w, v / 2, 4))


@grader("repeat_header")
def _(d, t, g):
    rows = _rows(_table(d, t))
    return bool(rows) and _on(rows[0].find(f"{W}trPr/{W}tblHeader"))


@grader("sort_table")
def _(d, t, g):
    rows = _rows(_table(d, t))
    if len(rows) < 3:
        return False
    head = [norm(_ctext(c)) for c in _cells(rows[0])]
    keys = []
    for m in re.finditer(r"by [“\"]([^”\"]+)[”\"] ?\(?(Ascending|Descending|in ascending order|"
                         r"in descending order)?", t, re.I):
        col = norm(m.group(1))
        if col not in head:
            return False
        keys.append((head.index(col), "desc" in (m.group(2) or "asc").casefold()))
    data = [[_ctext(c) for c in _cells(r)] for r in rows[1:]]
    data = [r for r in data if len(r) == len(head) and not norm(r[0]).startswith("total")]

    def conv(x):
        try:
            return (0, float(x.replace(",", "")))
        except ValueError:
            return (1, x.casefold())
    for a, b in zip(data, data[1:]):
        for idx, desc in keys:
            ka, kb = conv(a[idx]), conv(b[idx])
            if ka == kb:
                continue
            if (ka > kb) != desc:
                return False
            break
    return True


@grader("table_formula")
def _(d, t, g):
    rows = _rows(_table(d, t))
    for r in rows:
        cells = _cells(r)
        if cells and norm(_ctext(cells[0])).startswith("total"):
            fl = " ".join(fields(cells[-1])).upper().replace(" ", "")
            fmt = re.search(r"Use the (\S+) number format", t)
            return "=SUM(" in fl and (not fmt or fmt.group(1).rstrip(".") in fl)
    return False


def _alnum(s) -> str:
    return re.sub(r"[^0-9a-z]", "", norm(s))


def _tbl_style_is(d, tb, want) -> bool:
    """Word lưu "Grid Table 5 Dark - Accent 2" thành id GridTable5Dark-Accent2, tên "Grid Table 5 Dark Accent 2"."""
    sid = _attr(tb.find(f"{W}tblPr/{W}tblStyle"), "val")
    return bool(sid) and _alnum(want) in (_alnum(sid), _alnum(d.style_name(sid)))


def _look(tb, name):
    lk = tb.find(f"{W}tblPr/{W}tblLook")
    if lk is None:
        return None
    if _attr(lk, name) is not None:
        return _on(lk, name)
    bits = {"firstRow": 0x20, "lastRow": 0x40, "firstColumn": 0x80, "lastColumn": 0x100}
    try:
        return bool(int(_attr(lk, "val") or "0", 16) & bits[name])
    except ValueError:
        return None


@grader("table_style", "info_style")
def _(d, t, g):
    m = re.search(r"Apply the (.+?) table style", t, re.I)
    tb = _table(d, t)
    if tb is None or not m or not _tbl_style_is(d, tb, m.group(1)):
        return False
    if "first column" in t:
        return _look(tb, "firstColumn") is False
    if "Header Row" in t:
        return _look(tb, "firstRow") is False
    return True


@grader("info_noborder")
def _(d, t, g):
    tb = _table(d, t)
    if tb is None:
        return False
    for bd in tb.iter(W + "tblBorders", W + "tcBorders"):
        if any(_attr(x, "val") not in ("nil", "none") for x in bd):
            return False
    return tb.find(f"{W}tblPr/{W}tblBorders") is not None or not _tbl_style_is(d, tb, "Table Grid")


@grader("table_borders")
def _(d, t, g):
    m = re.search(r"apply a (.+?) pt (.+?) outside border", t)
    tb = _table(d, t)
    if tb is None or not m:
        return False
    sz, spec = _border_sz(t), color_spec(m.group(2))
    rows = _rows(tb)

    def good(el):
        return el is not None and _attr(el, "val") not in (None, "nil", "none") and \
            (sz is None or _near(_attr(el, "sz"), sz, 1)) and color_ok(el, spec, "border")
    tbd = tb.find(f"{W}tblPr/{W}tblBorders")
    for side in ("top", "left", "bottom", "right"):
        if tbd is not None and (good(tbd.find(W + side)) or (side in ("left", "right") and good(
                tbd.find(W + {"left": "start", "right": "end"}[side])))):
            continue
        if side == "top":
            edge = _cells(rows[0])
        elif side == "bottom":
            edge = _cells(rows[-1])
        elif side == "left":
            edge = [_cells(r)[0] for r in rows]
        else:
            edge = [_cells(r)[-1] for r in rows]
        if not all(good(c.find(f"{W}tcPr/{W}tcBorders/{W}{side}")) for c in edge):
            return False
    return True


@grader("table_shading")
def _(d, t, g):
    m = re.search(r"header row of the table to (.+?)\.?$", t)
    rows = _rows(_table(d, t))
    return bool(rows) and bool(m) and all(color_ok(c.find(f"{W}tcPr/{W}shd"), color_spec(m.group(1)), "fill")
                                          for c in _cells(rows[0]))


@grader("alt_table")
def _(d, t, g):
    tb = _table(d, t)
    want = norm(quotes(t)[0])
    if tb is None:
        return False
    if "title" in t.split("“")[0]:
        return norm(_attr(tb.find(f"{W}tblPr/{W}tblCaption"), "val")) == want
    return norm(_attr(tb.find(f"{W}tblPr/{W}tblDescription"), "val")) == want


def _list_paras(d, heading):
    return [p for p in section(d, heading) if p.tag == P and numbering(d, p)]


@grader("bullets_from_paras")
def _(d, t, g):
    m = re.search(r"the (\d+) paragraphs", t)
    ps = paras_from(d, sec_name(t), quotes(t)[1], int(m.group(1)))
    if len(ps) != int(m.group(1)):
        return False
    infos = [numbering(d, p) for p in ps]
    return all(i and i["numFmt"] == "bullet" for i in infos)


@grader("numbers_from_paras")
def _(d, t, g):
    m = re.search(r"the (\d+) paragraphs", t)
    f = re.search(r"uses the (.+?) number format", t)
    ps = paras_from(d, sec_name(t), quotes(t)[1], int(m.group(1)))
    if len(ps) != int(m.group(1)) or not f:
        return False
    kind, lvl_text = fmt_spec(f.group(1))
    infos = [numbering(d, p) for p in ps]
    return all(i and i["numFmt"] == kind and (i["lvlText"] or "").replace("%" + str(i["ilvl"] + 1), "%1") ==
               lvl_text for i in infos) and len({i["numId"] for i in infos}) == 1


@grader("num_format")
def _(d, t, g):
    f = re.search(r"number format of the numbered list to (.+?)\.?$", t)
    raw = f.group(1) if f else ""
    if raw.endswith("..") or raw.endswith(".)."):
        raw = raw[:-1]
    kind, lvl_text = fmt_spec(raw)
    ps = [p for p in _list_paras(d, sec_name(t))]
    infos = [numbering(d, p) for p in ps if numbering(d, p)["numFmt"] != "bullet"]
    return bool(infos) and all(i["numFmt"] == kind and (i["lvlText"] or "").replace(
        "%" + str(i["ilvl"] + 1), "%1") == lvl_text for i in infos if i["ilvl"] == 0)


@grader("custom_bullet")
def _(d, t, g):
    m = re.search(r"from the (.+?) font and character code [“\"](\w+)[”\"]", t)
    if not m:
        return None
    font, code = m.group(1), int(m.group(2), 16 if "Emoji" in m.group(1) else 10)
    chars = {chr(code), chr(0xF000 + code)} if code < 0x100 else {chr(code)}
    infos = [numbering(d, p) for p in _list_paras(d, sec_name(t))]
    return bool(infos) and all(i["numFmt"] == "bullet" and (i["lvlText"] or "") in chars and
                               norm(i["font"]) == norm(font) for i in infos)


@grader("list_level")
def _(d, t, g):
    m = re.search(r"to Level (\d)", t)
    item = quotes(t)[1]
    for p in section(d, sec_name(t)):
        if p.tag == P and norm(ptext(p)) == norm(item):
            info = numbering(d, p)
            return bool(info) and info["ilvl"] == int(m.group(1)) - 1
    return False


@grader("restart_num")
def _(d, t, g):
    m = re.search(r"starts at (\d+)", t)
    infos = [numbering(d, p) for p in _list_paras(d, sec_name(t))]
    infos = [i for i in infos if i["numFmt"] != "bullet" and i["ilvl"] == 0]
    if not infos:
        return False
    first = infos[0]
    return (first["startOverride"] or first["start"]) == m.group(1)


@grader("continue_num")
def _(d, t, g):
    infos = [numbering(d, p) for p in _list_paras(d, sec_name(t))]
    infos = [i for i in infos if i["numFmt"] != "bullet" and i["ilvl"] == 0]
    m = re.search(r"from (\d+) through (\d+)", t)
    if not infos or (m and len(infos) != int(m.group(2)) - int(m.group(1)) + 1):
        return False
    first = infos[0]
    for i in infos[1:]:
        if i["numId"] == first["numId"]:
            continue
        if i["abs"] != first["abs"] or i["startOverride"] is not None:
            return False
    return True


# ====================================================================== 4. Tham chiếu


def _note_part(kind):
    return f"word/{kind}s.xml"


@grader("cover_page")
def _(d, t, g):
    for gal in d.body.iter(W + "docPartGallery"):
        if norm(_attr(gal, "val")) == "cover pages":
            return True
    return False


@grader("footnote")
def _(d, t, g):
    kind = "endnote" if "endnote" in t.split(".")[0] else "footnote"
    head, note = quotes(t)[0], quotes(t)[-1]
    h = heading_para(d, head)
    if h is None or h.find(f".//w:{kind}Reference", NS) is None:
        return False
    root = d.xml(_note_part(kind))
    return root is not None and any(norm(note.rstrip(".")) in norm(ptext_all(n)) for n in root)


@grader("convert_notes")
def _(d, t, g):
    to_end = "footnotes to endnotes" in t
    src, dst = ("footnote", "endnote") if to_end else ("endnote", "footnote")
    return d.body.find(f".//w:{src}Reference", NS) is None and d.body.find(f".//w:{dst}Reference", NS) is not None


@grader("fn_format")
def _(d, t, g):
    fmt = "lowerRoman" if "i, ii" in t else "upperRoman" if "I, II" in t else \
        "lowerLetter" if "a, b" in t else "upperLetter" if "A, B" in t else "chicago" if "*" in t else None
    kind = "endnote" if "endnotes" in t else "footnote"
    for root in (settings(d), d.body):
        if root is None:
            continue
        for pr in root.iter(W + f"{kind}Pr"):
            if _attr(pr.find(W + "numFmt"), "val") == fmt:
                return True
    return False


@grader("toc_insert")
def _(d, t, g):
    return any(f.upper().startswith("TOC") and "\\c" not in f for f in fields(d.body))


@grader("custom_toc")
def _(d, t, g):
    m = re.search(r"shows (\d) heading levels", t)
    s = re.search(r"formatted with the (.+?) style at TOC level (\d)", t)
    for f in fields(d.body):
        if f.upper().startswith("TOC"):
            ok = not m or re.search(rf'\\o\s*"1-{m.group(1)}"', f)
            ok2 = not s or re.search(rf'\\t\s*"[^"]*{re.escape(s.group(1))}\s*[,;]\s*{s.group(2)}', f, re.I)
            if ok and ok2:
                return True
    return False


@grader("toc_modify")
def _(d, t, g):
    n = "2" if "Heading 2" in t else "1"
    return any(f.upper().startswith("TOC") and re.search(rf'\\o\s*"1-{n}"', f) for f in fields(d.body))


@grader("table_of_figures")
def _(d, t, g):
    return any(f.upper().startswith("TOC") and re.search(r'\\c\s*"Table"', f, re.I) for f in fields(d.body))


@grader("caption_pic", "caption_table")
def _(d, t, g):
    label = "Table" if "Table label" in t else "Figure"
    q = quotes(t)
    if label == "Table":
        want = norm(f"Table 1{q[-1]}")
    else:
        want = norm(q[-1])
    for b in section(d, sec_name(t)):
        if b.tag != P or norm(ptext(b)) != want:
            continue
        if not any(f.upper().startswith(f"SEQ {label.upper()}") for f in fields(b)):
            continue
        if label == "Table":
            where = "above" if "above the table" in t else "below"
            nb = b.getnext() if where == "above" else b.getprevious()
            while nb is not None and nb.tag == P and not ptext(nb).strip():
                nb = nb.getnext() if where == "above" else nb.getprevious()
            return nb is not None and nb.tag == TBL
        return True
    return False


@grader("cross_ref")
def _(d, t, g):
    for p in sec_paras(d, sec_name(t)):
        for f in fields(p):
            if re.match(r"REF\s+_Ref", f) and "\\p" not in f:
                return True
    return False


@grader("citation")
def _(d, t, g):
    tag = quotes(t)[-1]
    idx = 1 if "first paragraph" in t else 2 if "second paragraph" in t else None
    ps = sec_paras(d, sec_name(t))
    cands = [ps[idx - 1]] if idx and len(ps) >= idx else ps
    return any(any(re.match(rf"CITATION\s+{re.escape(tag)}\b", f) for f in fields(p)) for p in cands)


def _sources_xml(d) -> str:
    return "\n".join(d.raw(n) for n in d.glob(r"customXml/item\d+\.xml") if "Sources" in d.raw(n))


@grader("citation_source")
def _(d, t, g):
    m = re.search(r"Title: (.+?);", t)
    src = _sources_xml(d)
    if not m or norm(m.group(1)) not in norm(re.sub(r"<[^>]+>", " ", src)):
        return False
    ps = sec_paras(d, sec_name(t))
    return any(any(f.upper().startswith("CITATION") for f in fields(p)) for p in ps)


@grader("biblio_style")
def _(d, t, g):
    m = re.search(r"style of the document to (\w+)", t)
    src = _sources_xml(d)
    return bool(m) and bool(re.search(rf'(StyleName="{m.group(1)}|SelectedStyle="[^"]*{m.group(1)})', src, re.I))


@grader("index")
def _(d, t, g):
    phrase = quotes(t)[0]
    xe = [f for f in fields(d.body) if re.match(rf'XE\s+"{re.escape(phrase)}"', f)]
    idx = [f for f in fields(d.body) if f.upper().startswith("INDEX")]
    m = re.search(r"and (\w+) columns", t)
    ncol = {"one": 1, "two": 2, "three": 3}.get(m.group(1) if m else "", None)
    if not xe or not idx:
        return False
    return ncol is None or any(re.search(rf'\\c\s*"{ncol}"', f) for f in idx)


# ====================================================================== 5. Đồ họa


def _pic(d, t):
    drs = sec_objects(d, t, "pic")
    return drs[0] if drs else None


@grader("pic_size")
def _(d, t, g):
    dr = _pic(d, t)
    if dr is None:
        return False
    cx, cy = extent(dr)
    v = inches(t)[0] * EMU_IN
    size_ok = _near(cx if "width" in t else cy, v, 0.03 * EMU_IN)
    if g.get("cx") and g.get("cy") and cx and cy:
        return size_ok and abs(cx / cy - g["cx"] / g["cy"]) < 0.03 * (g["cx"] / g["cy"])
    return size_ok


@grader("chart_size", "sa_size")
def _(d, t, g):
    kind = "chart" if t.find("chart") >= 0 and "SmartArt" not in t else "dgm"
    drs = sec_objects(d, t, kind)
    if not drs:
        return False
    cx, cy = extent(drs[0])
    h = re.search(r"(\d+(?:\.\d+)?)\" \([^)]*\) high", t)
    w = re.search(r"(\d+(?:\.\d+)?)\" \([^)]*\) wide", t)
    return bool(h and w) and _near(cy, float(h.group(1)) * EMU_IN, 0.03 * EMU_IN) and \
        _near(cx, float(w.group(1)) * EMU_IN, 0.03 * EMU_IN)


def _sppr(dr):
    return dr.find(".//pic:spPr", NS) if dr.find(".//pic:spPr", NS) is not None else dr.find(".//wps:spPr", NS)


EFFECTS = {"reflection": "reflection", "soft edge": "softEdge", "glow": "glow", "shadow": ("outerShdw", "innerShdw",
                                                                                           "prstShdw"),
           "bevel": ("bevelT",), "3-d rotation": ("scene3d",)}


def _has_effect(root, desc) -> bool:
    low = desc.casefold()
    names = {etree.QName(x).localname for x in root.iter() if isinstance(x.tag, str)}
    for key, tags in EFFECTS.items():
        if key in low:
            tags = (tags,) if isinstance(tags, str) else tags
            if not names & set(tags):
                return False
            if key == "soft edge":
                m = re.search(r"(\d+(?:\.\d+)?) point soft", low)
                se = root.find(".//a:softEdge", NS)
                if m and se is not None and not _near(_int(se.get("rad")), float(m.group(1)) * 12700, 13000):
                    return False
            if key == "glow":
                m = re.search(r"glow: (\d+) point", low)
                gl = root.find(".//a:glow", NS)
                if m and gl is not None and not _near(_int(gl.get("rad")), int(m.group(1)) * 12700, 13000):
                    return False
            return True
    if "round" in low and "bevel" in low:
        return "bevelT" in names
    return False


@grader("pic_effect")
def _(d, t, g):
    m = re.search(r"apply the (.+?) effect to the picture", t, re.I)
    dr = _pic(d, t)
    return dr is not None and bool(m) and _has_effect(_sppr(dr), m.group(1))


PIC_STYLES = {
    "soft edge rectangle": lambda sp: sp.find(".//a:softEdge", NS) is not None,
    "center shadow rectangle": lambda sp: sp.find(".//a:outerShdw", NS) is not None,
    "drop shadow rectangle": lambda sp: sp.find(".//a:outerShdw", NS) is not None,
    "metal oval": lambda sp: _geom(sp) == "ellipse",
    "rounded diagonal corner, white": lambda sp: _geom(sp) == "round2DiagRect",
    "reflected rounded rectangle": lambda sp: _geom(sp) == "roundRect" and sp.find(".//a:reflection", NS)
    is not None,
    "simple frame, black": lambda sp: sp.find("a:ln", NS) is not None and _int(sp.find("a:ln", NS).get("w"), 0) > 0,
    "simple frame, white": lambda sp: sp.find("a:ln", NS) is not None and _int(sp.find("a:ln", NS).get("w"), 0) > 0,
    "beveled matte, white": lambda sp: sp.find(".//a:bevelT", NS) is not None or _geom(sp) == "roundRect",
}


def _geom(sp):
    g = sp.find("a:prstGeom", NS) if sp is not None else None
    return g.get("prst") if g is not None else None


@grader("pic_style")
def _(d, t, g):
    m = re.search(r"apply the (.+?) picture style", t, re.I)
    dr = _pic(d, t)
    fn = PIC_STYLES.get(norm(m.group(1))) if m else None
    if fn is None or dr is None:
        return None if fn is None else False
    return bool(fn(_sppr(dr)))


ARTISTIC = {"marker": "artisticMarker", "glow diffused": "artisticGlowDiffused"}


@grader("art_effect")
def _(d, t, g):
    m = re.search(r"apply the (.+?) artistic effect", t, re.I)
    dr = _pic(d, t)
    if dr is None or not m:
        return False
    want = ARTISTIC.get(norm(m.group(1)), "artistic" + m.group(1).title().replace(" ", ""))
    names = {etree.QName(x).localname.casefold() for x in dr.iter() if isinstance(x.tag, str)}
    return any(n.startswith(want.casefold()) for n in names)


@grader("pic_remove_bg")
def _(d, t, g):
    dr = _pic(d, t)
    return dr is not None and any(isinstance(x.tag, str) and etree.QName(x).localname == "backgroundRemoval"
                                  for x in dr.iter())


SHAPES = {"oval": "ellipse", "flowchart: alternate process": "flowChartAlternateProcess",
          "rectangle: rounded corners": "roundRect", "cloud": "cloud", "ribbon: tilted up": "ribbon2",
          "ribbon: tilted down": "ribbon", "star: 5 points": "star5", "scroll: horizontal": "horizontalScroll",
          "explosion: 8 points": "irregularSeal1", "speech bubble: rectangle": "wedgeRectCallout",
          "heart": "heart", "rectangle": "rect", "isosceles triangle": "triangle", "hexagon": "hexagon",
          "rectangle: snip single corner": "snip1Rect", "flowchart: process": "flowChartProcess",
          "flowchart: document": "flowChartDocument", "teardrop": "teardrop"}


@grader("pic_crop_shape")
def _(d, t, g):
    m = re.search(r"to the (.+?) shape", t)
    dr = _pic(d, t)
    want = SHAPES.get(norm(m.group(1))) if m else None
    if want is None:
        return None
    return dr is not None and _geom(_sppr(dr)) == want


@grader("pic_border")
def _(d, t, g):
    m = re.search(r"Add a (.+?) pt (.+?) border to the picture", t)
    dr = _pic(d, t)
    if dr is None or not m:
        return False
    ln = _sppr(dr).find("a:ln", NS)
    if ln is None or ln.find("a:noFill", NS) is not None:
        return False
    w = {"½": 0.5, "¾": 0.75, "1": 1, "1 ½": 1.5, "2 ¼": 2.25, "3": 3, "4 ½": 4.5, "6": 6}.get(m.group(1).strip())
    return (w is None or _near(_int(ln.get("w")), w * 12700, 2000)) and dml_color_ok(ln, color_spec(m.group(2)))


def _position_ok(dr, v, h):
    if dr is None or dr.tag != "{%s}anchor" % NS["wp"]:
        return False
    ph, pv = dr.find("wp:positionH", NS), dr.find("wp:positionV", NS)
    ah = ph.find("wp:align", NS) if ph is not None else None
    av = pv.find("wp:align", NS) if pv is not None else None
    return ah is not None and av is not None and ah.text == h and av.text == v


def _wrap(dr, kind):
    if dr is None:
        return False
    if kind == "inline":
        return dr.tag == "{%s}inline" % NS["wp"]
    if kind == "behind":
        return dr.get("behindDoc") == "1" and dr.find("wp:wrapNone", NS) is not None
    if kind == "front":
        return dr.get("behindDoc") in ("0", None) and dr.find("wp:wrapNone", NS) is not None
    return dr.find(f"wp:wrap{kind}", NS) is not None


WRAP = {"square": "Square", "tight": "Tight", "through": "Through", "top and bottom": "TopAndBottom",
        "behind text": "behind", "in front of text": "front", "in line with text": "inline"}


def _pos_text(t):
    m = re.search(r"in the (Top|Middle|Bottom) (Left|Center|Right) of the page", t, re.I) or \
        re.search(r"at the (top|middle|bottom) (left|center|right) of the page", t, re.I)
    if not m:
        return None
    return {"top": "top", "middle": "center", "bottom": "bottom"}[m.group(1).lower()], m.group(2).lower()


@grader("pic_position")
def _(d, t, g):
    dr = _pic(d, t)
    pos = _pos_text(t)
    wm = re.search(r"with (\w+(?: and \w+)?) text wrapping", t)
    return pos is not None and _position_ok(dr, *pos) and (not wm or _wrap(dr, WRAP[wm.group(1).lower()]))


@grader("pic_wrap")
def _(d, t, g):
    m = re.search(r"text wrapping for the picture to (.+?)\.?$", t)
    return bool(m) and _wrap(_pic(d, t), WRAP.get(m.group(1).lower(), m.group(1)))


@grader("pic_change")
def _(d, t, g):
    dr = _pic(d, t)
    if dr is None:
        return False
    img = md5(pic_image(d, dr))
    if g.get("img") and img == g["img"]:
        return False
    cx, cy = extent(dr)
    return not g.get("cx") or (_near(cx, g["cx"], 0.05 * g["cx"]) and _near(cy, g["cy"], 0.05 * g["cy"]))


@grader("pic_insert")
def _(d, t, g):
    return len(sec_objects(d, t, "pic")) > g.get("n", 0)


@grader("model3d")
def _(d, t, g):
    drs = sec_objects(d, t, "model3d")
    if not drs:
        drs = [dr for dr in sec_objects(d, t, None) if "model3d" in etree.tostring(dr).decode("utf-8", "ignore")]
    if "In Line with Text" in t:
        return any(dr.tag == "{%s}inline" % NS["wp"] for dr in drs)
    return bool(drs)


@grader("shape_insert")
def _(d, t, g):
    m = re.search(r"insert an? (.+?) shape that contains the text [“\"]([^”\"]+)", t)
    if not m:
        return None
    want = SHAPES.get(norm(m.group(1)))
    for dr in drawings([d.body], "wps"):
        if _geom(dr.find(".//wps:spPr", NS)) == want and norm(m.group(2).rstrip(".!")) in norm(txbx_text(dr)):
            pos = _pos_text(t)
            wm = re.search(r"with (\w+) text wrapping", t)
            return (pos is None or _position_ok(dr, *pos)) and (not wm or _wrap(dr, WRAP[wm.group(1).lower()]))
    return False


@grader("textbox_type")
def _(d, t, g):
    val = norm(quotes(t)[-1])
    return any(val in norm(txbx_text(dr)) for dr in sec_objects(d, t, "wps"))


@grader("callout_text")
def _(d, t, g):
    start, val = quotes(t)[1], quotes(t)[2]
    dr = textbox_starting(d, start)
    if dr is None:
        return False
    paras = [norm(ptext(p)) for p in dr.find(".//w:txbxContent", NS).iter(P)]
    return norm(val) in paras


@grader("callout_pos")
def _(d, t, g):
    dr = textbox_starting(d, quotes(t)[0])
    pos = _pos_text(t)
    wm = re.search(r"with (\w+) text wrapping", t)
    return pos is not None and _position_ok(dr, *pos) and (not wm or _wrap(dr, WRAP[wm.group(1).lower()]))


@grader("shape_style")
def _(d, t, g):
    m = re.search(r"apply the (.+?) shape style", t, re.I)
    dr = textbox_starting(d, quotes(t)[-1])
    if dr is None or not m:
        return False
    acc = re.search(r"Accent (\d)", m.group(1))
    style = dr.find(".//wps:style", NS)
    if style is None:
        return False
    need = f"accent{acc.group(1)}" if acc else None
    has = [c.get("val") for c in style.iter("{%s}schemeClr" % NS["a"])]
    return need is None or need in has


@grader("link_textbox")
def _(d, t, g):
    return d.body.find(".//wps:linkedTxbx", NS) is not None or "linkedTxbx" in d.raw("word/document.xml")


@grader("wordart")
def _(d, t, g):
    title = g.get("title")
    for dr in textboxes(d):
        if title and norm(title) not in norm(txbx_text(dr)):
            continue
        tb = dr.find(".//w:txbxContent", NS)
        if any(_w14_effects(d, p, "") for p in tb.iter(P) if ptext(p).strip()):
            return True
    return False


# ---------- biểu đồ


def _chart(d, t):
    parts = chart_parts(d, sec_objects(d, t, "chart"))
    return (parts[0], d.xml(parts[0])) if parts else (None, None)


def _chart_fp(d, part):
    rels = d.rels(part)
    blobs = [d.files.get(p, b"") for p in rels.values() if "style" in p or "colors" in p]
    root = d.xml(part)
    for tag in ("style",):
        for el in root.iter("{%s}%s" % (NS["c"], tag)):
            blobs.append(etree.tostring(el))
    for el in root.iter("{%s}spPr" % NS["c"]):
        blobs.append(etree.tostring(el))
    lg = root.find(".//c:legend", NS)
    blobs.append(etree.tostring(lg) if lg is not None else b"")
    blobs.append(b"title" if root.find(".//c:chart/c:title", NS) is not None else b"")
    return md5(b"".join(blobs))


def _chart_value(d, part, t):
    root = d.xml(part)
    q = quotes(t)
    if len(q) < 2:
        return None
    a, b = norm(q[0]), norm(q[1])
    for ser in root.iter("{%s}ser" % NS["c"]):
        name = norm(" ".join(x.text or "" for x in ser.find("c:tx", NS).iter("{%s}v" % NS["c"]))) \
            if ser.find("c:tx", NS) is not None else ""
        cats = [norm(x.text) for x in ser.find("c:cat", NS).iter("{%s}v" % NS["c"])] \
            if ser.find("c:cat", NS) is not None else []
        vals = ser.find("c:val", NS)
        pts = {int(p.get("idx")): p.findtext("c:v", namespaces=NS) for p in vals.iter("{%s}pt" % NS["c"])} \
            if vals is not None else {}
        for s, c in ((a, b), (b, a)):
            if name == s and c in cats:
                v = pts.get(cats.index(c))
                return float(v) if v not in (None, "") else None
    return None


@grader("chart_data")
def _(d, t, g):
    part, root = _chart(d, t)
    m = re.search(r"Change the value to (\d+(?:\.\d+)?)", t)
    if part is None or not m:
        return False
    v = _chart_value(d, part, t)
    if v is not None:
        return abs(v - float(m.group(1))) < 1e-6
    return any(_near(float(x.text or "nan"), float(m.group(1)), 1e-6) for x in root.iter("{%s}v" % NS["c"])
               if re.fullmatch(r"-?\d+(\.\d+)?", x.text or ""))


CHART_TYPES = {
    "clustered column": ("barChart", "col", "clustered"), "clustered bar": ("barChart", "bar", "clustered"),
    "stacked column": ("barChart", "col", "stacked"), "stacked bar": ("barChart", "bar", "stacked"),
    "100% stacked column": ("barChart", "col", "percentStacked"), "line": ("lineChart", None, None),
    "line with markers": ("lineChart", None, "markers"), "pie": ("pieChart", None, None),
    "doughnut": ("doughnutChart", None, None), "area": ("areaChart", None, None),
    "3-d pie": ("pie3DChart", None, None), "3-d clustered column": ("bar3DChart", "col", None),
}


def _chart_type_ok(root, name):
    spec = CHART_TYPES.get(norm(name))
    if spec is None or root is None:
        return None
    tag, direction, grouping = spec
    el = root.find(f".//c:plotArea/c:{tag}", NS)
    if el is None:
        return False
    if direction and (el.find("c:barDir", NS) is None or el.find("c:barDir", NS).get("val") != direction):
        return False
    if grouping == "markers":
        mk = el.find("c:marker", NS)
        sers = el.findall("c:ser", NS)
        return not any(s.find("c:marker/c:symbol", NS) is not None and s.find("c:marker/c:symbol", NS).get(
            "val") == "none" for s in sers) and (mk is None or mk.get("val") in ("1", "true", None) or bool(sers))
    if grouping:
        gr = el.find("c:grouping", NS)
        if (gr.get("val") if gr is not None else "clustered") != grouping:
            return False
    return True


@grader("chart_type")
def _(d, t, g):
    m = re.search(r"change the chart to an? (.+?) chart", t)
    part, root = _chart(d, t)
    return bool(m) and root is not None and _chart_type_ok(root, m.group(1))


@grader("chart_insert")
def _(d, t, g):
    m = re.search(r"insert an? (.+?) chart", t)
    drs = sec_objects(d, t, "chart")
    if len(drs) <= g.get("n", 0) or not m:
        return False
    series = quotes(t)[-1]
    nums = [float(x) for x in re.findall(r"\(([\d.]+)\)", t)]
    for part in chart_parts(d, drs):
        root = d.xml(part)
        if not _chart_type_ok(root, m.group(1)):
            continue
        for ser in root.iter("{%s}ser" % NS["c"]):
            name = " ".join(x.text or "" for x in ser.find("c:tx", NS).iter("{%s}v" % NS["c"])) \
                if ser.find("c:tx", NS) is not None else ""
            vals = [float(p.findtext("c:v", namespaces=NS)) for p in ser.find("c:val", NS).iter("{%s}pt" % NS["c"])] \
                if ser.find("c:val", NS) is not None else []
            if norm(name) == norm(series) and (not nums or [round(v, 4) for v in vals] == [round(v, 4) for v in nums]):
                return True
    return False


@grader("chart_title")
def _(d, t, g):
    part, root = _chart(d, t)
    want = quotes(t)[-1]
    title = root.find(".//c:chart/c:title", NS) if root is not None else None
    return title is not None and norm("".join(x.text or "" for x in title.iter("{%s}t" % NS["a"]))) == norm(want)


@grader("chart_axis")
def _(d, t, g):
    part, root = _chart(d, t)
    want = quotes(t)[-1]
    if root is None:
        return False
    tag = "valAx" if "vertical" in t else "catAx"
    for ax in root.iter("{%s}%s" % (NS["c"], tag)):
        title = ax.find("c:title", NS)
        if title is not None and norm("".join(x.text or "" for x in title.iter("{%s}t" % NS["a"]))) == norm(want):
            return True
    return False


@grader("chart_gridlines")
def _(d, t, g):
    part, root = _chart(d, t)
    return root is not None and root.find(".//c:majorGridlines", NS) is None


@grader("chart_legend")
def _(d, t, g):
    part, root = _chart(d, t)
    return root is not None and root.find(".//c:chart/c:legend", NS) is None


@grader("chart_labels")
def _(d, t, g):
    part, root = _chart(d, t)
    if root is None:
        return False
    m = re.search(r"Position the labels (.+?)\.?$", t)
    pos = {"outside end": "outEnd", "inside end": "inEnd", "center": "ctr", "inside base": "inBase",
           "best fit": "bestFit"}.get(norm(m.group(1)) if m else "", None)
    for dl in root.iter("{%s}dLbls" % NS["c"]):
        sv = dl.find("c:showVal", NS)
        if sv is None or sv.get("val") not in ("1", "true"):
            continue
        p = dl.find("c:dLblPos", NS)
        if pos is None or (p is not None and p.get("val") == pos):
            return True
    return False


@grader("chart_switch")
def _(d, t, g):
    part, root = _chart(d, t)
    if root is None:
        return False
    ser = root.find(".//c:ser", NS)
    cat = ser.find("c:cat", NS) if ser is not None else None
    vals = [x.text or "" for x in cat.iter("{%s}v" % NS["c"])] if cat is not None else []
    return bool(vals) and all(re.fullmatch(r"(19|20)\d\d", v.strip()) for v in vals)


@grader("chart_alt")
def _(d, t, g):
    val = norm(quotes(t)[0])
    return any(norm(docpr(dr).get("descr")) == val for dr in sec_objects(d, t, "chart"))


@grader("chart_style", "chart_colors", "chart_layout")
def _(d, t, g):
    part, root = _chart(d, t)
    if root is None:
        return False
    if "Layout 3" in t or "Quick Layout" in t:
        lg = root.find(".//c:chart/c:legend/c:legendPos", NS)
        m = re.search(r"Layout (\d+)", t)
        if m and m.group(1) == "3":
            return root.find(".//c:chart/c:title", NS) is not None and lg is not None and lg.get("val") == "b"
    if "colors of the chart" in t:
        colors = [p for p in d.rels(part).values() if "colors" in p]
        if colors:
            raw = d.raw(colors[0])
            if "Monochromatic" in t:
                return 'meth="withinLinear"' in raw or bool(g.get("fp")) and _chart_fp(d, part) != g["fp"]
    return bool(g.get("fp")) and _chart_fp(d, part) != g["fp"]


# ---------- SmartArt


def _dgm(d, t):
    drs = sec_objects(d, t, "dgm")
    return drs[0] if drs else None


def _unique_id(d, part):
    root = d.xml(part) if part else None
    return root.get("uniqueId", "") if root is not None else ""


def _dgm_texts(d, dr) -> list[str]:
    part = dgm_parts(d, dr).get("dm")
    root = d.xml(part) if part else None
    if root is None:
        return []
    out = []
    for pt in root.iter("{%s}pt" % NS["dgm"]):
        if pt.get("type") in (None, "node"):
            txt = "".join(x.text or "" for x in pt.iter("{%s}t" % NS["a"]))
            if txt.strip():
                out.append(txt)
    return out


SA_LAYOUTS = {"basic block list": "default", "basic process": "process1", "basic cycle": "cycle2",
              "vertical bullet list": "vList2", "basic chevron process": "chevron1",
              "continuous block process": "hProcess9", "basic timeline": "hProcess11", "radial cycle": "radial6",
              "alternating hexagons": "AlternatingHexagons", "vertical box list": "vList5",
              "basic venn": "venn1", "step up process": "StepUpProcess", "segmented cycle": "cycle8"}
SA_STYLES = {"simple fill": "simple1", "white outline": "simple2", "subtle effect": "simple3",
             "moderate effect": "simple4", "intense effect": "simple5", "polished": "3d1", "inset": "3d2",
             "cartoon": "3d3", "powder": "3d4", "brick scene": "3d5", "flat scene": "3d6",
             "metallic scene": "3d7", "sunset scene": "3d8", "bird's eye scene": "3d9"}


def _sa_color_id(name):
    low = norm(name)
    m = re.match(r"colorful range - accent colors (\d) to (\d)", low)
    if low == "colorful - accent colors":
        return "colorful1"
    if m:
        return f"colorful{int(m.group(1))}"
    m = re.match(r"(colored fill|colored outline|gradient range|gradient loop|transparent gradient range)"
                 r" - accent (\d)", low)
    if m:
        n = {"colored outline": 1, "colored fill": 2, "transparent gradient range": 3, "gradient range": 4,
             "gradient loop": 5}[m.group(1)]
        return f"accent{m.group(2)}_{n}"
    return None


def _id_is(uid, short):
    return bool(uid) and uid.rstrip("#").rsplit("/", 1)[-1].casefold() == short.casefold()


@grader("sa_layout")
def _(d, t, g):
    m = re.search(r"SmartArt graphic in the .+? section to (.+?)\.?$", t)
    dr = _dgm(d, t)
    want = SA_LAYOUTS.get(norm(m.group(1))) if m else None
    if want is None:
        return None
    return dr is not None and _id_is(_unique_id(d, dgm_parts(d, dr).get("lo")), want)


@grader("sa_style")
def _(d, t, g):
    m = re.search(r"Apply the (.+?) SmartArt style", t, re.I)
    dr = _dgm(d, t)
    want = SA_STYLES.get(norm(m.group(1))) if m else None
    if dr is None or want is None:
        return None if want is None else False
    return _id_is(_unique_id(d, dgm_parts(d, dr).get("qs")), want)


@grader("sa_colors")
def _(d, t, g):
    m = re.search(r"section to (.+?)\.?$", t)
    dr = _dgm(d, t)
    want = _sa_color_id(m.group(1)) if m else None
    if dr is None:
        return False
    uid = _unique_id(d, dgm_parts(d, dr).get("cs"))
    if want:
        return _id_is(uid, want)
    return bool(g.get("cs")) and uid != g["cs"]


@grader("sa_alt")
def _(d, t, g):
    dr = _dgm(d, t)
    return dr is not None and norm(docpr(dr).get("descr")) == norm(quotes(t)[0])


@grader("sa_text")
def _(d, t, g):
    dr = _dgm(d, t)
    new = quotes(t)[-1]
    return dr is not None and norm(new) in [norm(x) for x in _dgm_texts(d, dr)]


@grader("sa_add")
def _(d, t, g):
    q = quotes(t)
    after, new = q[1], q[2]
    dr = _dgm(d, t)
    if dr is None:
        return False
    texts = [norm(x) for x in _dgm_texts(d, dr)]
    if norm(new) not in texts:
        return False
    order = _dgm_order(d, dr)
    if order and norm(after) in order and norm(new) in order:
        return order.index(norm(new)) == order.index(norm(after)) + 1
    return True


def _dgm_model(d, dr):
    part = dgm_parts(d, dr).get("dm")
    root = d.xml(part) if part else None
    if root is None:
        return None, {}, []
    pts = {pt.get("modelId"): pt for pt in root.iter("{%s}pt" % NS["dgm"])}
    cxns = [c for c in root.iter("{%s}cxn" % NS["dgm"]) if c.get("type") in (None, "parOf")]
    return root, pts, cxns


def _dgm_order(d, dr) -> list[str]:
    """Thứ tự các mục cấp 1 (con của điểm gốc 'doc')."""
    root, pts, cxns = _dgm_model(d, dr)
    if root is None:
        return []
    doc_id = next((k for k, v in pts.items() if v.get("type") == "doc"), None)
    kids = sorted((c for c in cxns if c.get("srcId") == doc_id), key=lambda c: _int(c.get("srcOrd"), 0))
    out = []
    for c in kids:
        pt = pts.get(c.get("destId"))
        if pt is not None:
            out.append(norm("".join(x.text or "" for x in pt.iter("{%s}t" % NS["a"]))))
    return out


@grader("sa_demote")
def _(d, t, g):
    q = quotes(t)
    child, parent = q[1], q[2]
    dr = _dgm(d, t)
    if dr is None:
        return False
    root, pts, cxns = _dgm_model(d, dr)
    ids = {norm("".join(x.text or "" for x in pt.iter("{%s}t" % NS["a"]))): k for k, pt in pts.items()
           if pt.get("type") in (None, "node")}
    c_id, p_id = ids.get(norm(child)), ids.get(norm(parent))
    return bool(c_id and p_id) and any(c.get("srcId") == p_id and c.get("destId") == c_id for c in cxns)


@grader("sa_rtl")
def _(d, t, g):
    dr = _dgm(d, t)
    part = dgm_parts(d, dr).get("dm") if dr is not None else None
    return bool(part) and bool(re.search(r'<(?:\w+:)?dir\s+val="rev"', d.raw(part)))


@grader("sa_effect")
def _(d, t, g):
    m = re.search(r"apply the (.+?) shape effect", t, re.I)
    dr = _dgm(d, t)
    if dr is None or not m:
        return False
    parts = dgm_parts(d, dr)
    roots = [d.xml(parts.get("dm"))] if parts.get("dm") else []
    dwg = [p for p in d.rels(parts["dm"]).values() if "drawing" in p] if parts.get("dm") else []
    roots += [d.xml(p) for p in dwg]
    desc = m.group(1)
    if "soft round" in desc.casefold():
        return any(r is not None and any(b.get("prst") == "softRound" for b in r.iter("{%s}bevelT" % NS["a"]))
                   for r in roots)
    return any(r is not None and _has_effect(r, desc) for r in roots)


@grader("smartart")
def _(d, t, g):
    m = re.search(r"insert an? (.+?) SmartArt graphic", t)
    st = re.search(r"Apply the (.+?) SmartArt style", t, re.I)
    items = [norm(x) for x in quotes(t)[1:]]
    for dr in sec_objects(d, t, "dgm"):
        parts = dgm_parts(d, dr)
        want = SA_LAYOUTS.get(norm(m.group(1))) if m else None
        if want and not _id_is(_unique_id(d, parts.get("lo")), want):
            continue
        order = _dgm_order(d, dr) or [norm(x) for x in _dgm_texts(d, dr)]
        if order != items:
            continue
        if st and SA_STYLES.get(norm(st.group(1))) and not _id_is(_unique_id(d, parts.get("qs")),
                                                                  SA_STYLES[norm(st.group(1))]):
            continue
        return True
    return False


# ====================================================================== 6. Cộng tác


@grader("comment_add")
def _(d, t, g):
    q = quotes(t)
    phrase, body = q[1], q[2]
    cm = comments(d)
    ranges = comment_ranges(d)
    return any(norm(body.rstrip(".")) in norm(txt) and norm(phrase) in norm(ranges.get(cid, ""))
               for cid, txt in cm.items())


@grader("comment_delete")
def _(d, t, g):
    phrase = quotes(t)[-1]
    ranges = comment_ranges(d)
    if any(norm(phrase) in norm(txt) for txt in ranges.values()):
        return False
    return not g.get("n") or len(comments(d)) == g["n"] - 1


@grader("comment_delete_all")
def _(d, t, g):
    return not comments(d) and d.body.find(".//w:commentReference", NS) is None


@grader("comment_reply")
def _(d, t, g):
    reply = quotes(t)[-1]
    cm = comments(d)
    if not any(norm(reply.rstrip(".")) == norm(x).rstrip(".") for x in cm.values()):
        return False
    ext = d.xml("word/commentsExtended.xml")
    if ext is not None:
        return any(_attr(c, "paraIdParent", "w15") for c in ext.iter())
    return len(cm) > g.get("n", 0)


@grader("comment_resolve")
def _(d, t, g):
    ext = d.xml("word/commentsExtended.xml")
    if ext is None:
        return False
    return sum(1 for c in ext.iter() if _attr(c, "done", "w15") == "1") > g.get("done", 0)


@grader("track_on")
def _(d, t, g):
    old, new = quotes(t)[0], quotes(t)[1]
    rev = revisions(d)
    return _track_on(d) and any(new in x for x in rev["ins"]) and any(old in x for x in rev["del"])


@grader("track_accept")
def _(d, t, g):
    rev = revisions(d)
    if rev["ins"] or rev["del"] or rev["fmt"]:
        return False
    body = doc_text(d)
    ins_ok = all(x.strip() in body for x in g.get("ins", []))
    del_ok = all(x.strip() not in body for x in g.get("del", []) if len(x.strip()) > 3)
    return ins_ok and del_ok


@grader("track_reject")
def _(d, t, g):
    keep = quotes(t)[0]
    rev = revisions(d)
    if rev["ins"] or rev["del"] or rev["fmt"]:
        return False
    body = doc_text(d)
    if keep not in body:
        return False
    others = [x for x in g.get("ins", []) if norm(x) != norm(keep) and len(x.strip()) > 3]
    return all(x.strip() not in body for x in others) and all(x.strip() in body for x in g.get("del", []))


# ====================================================================== Expert: lưu ra ngoài file


def _appdata() -> Path:
    return Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming")


@grader("quickpart")
def _(d, t, g):
    name = quotes(t)[-1]
    cands = list((_appdata() / "Microsoft" / "Document Building Blocks").glob("**/*.dotx"))
    cands += list((_appdata() / "Microsoft" / "Templates").glob("Normal.dotm"))
    for f in cands:
        try:
            with zipfile.ZipFile(f) as z:
                raw = "".join(z.read(n).decode("utf-8", "ignore") for n in z.namelist() if "glossary" in n)
        except (OSError, zipfile.BadZipFile):
            continue
        if re.search(r'<w:name w:val="' + re.escape(name) + '"', raw):
            return True
    return False


@grader("custom_styleset")
def _(d, t, g):
    name = quotes(t)[0]
    f = _appdata() / "Microsoft" / "QuickStyles" / f"{name}.dotx"
    return f.is_file()


@grader("macro")
def _(d, t, g):
    stem = d.path.stem
    f = saved_copy(d.path, stem, [".docm"])
    if f is None:
        return False
    try:
        with zipfile.ZipFile(f) as z:
            return any(n.endswith("vbaProject.bin") for n in z.namelist())
    except zipfile.BadZipFile:
        return False


@grader("mailmerge")
def _(d, t, g):
    st = settings(d)
    mm = st.find(W + "mailMerge") if st is not None else None
    if mm is None:
        return False
    src = quotes(t)[0] if quotes(t) else ""
    m = re.search(r"Use the (\S+\.xlsx)", t)
    raw = etree.tostring(mm).decode("utf-8", "ignore") + d.raw("word/_rels/settings.xml.rels")
    if m and m.group(1).casefold() not in raw.casefold():
        return False
    field = re.search(r"insert the (\w+) merge field", t)
    p = find_para(d, src or "Dear", start=True)
    return p is not None and (not field or any(re.match(rf"MERGEFIELD\s+\"?{field.group(1)}\b", f)
                                               for f in fields(p)))


@grader("content_control")
def _(d, t, g):
    title = quotes(t)[-1]
    for sdt in d.body.iter(W + "sdt"):
        pr = sdt.find(W + "sdtPr")
        if pr is not None and norm(_attr(pr.find(W + "alias"), "val")) == norm(title):
            if "Plain Text" in t:
                return pr.find(W + "text") is not None
            return True
    return False


@grader("fields")
def _(d, t, g):
    m = re.search(r"insert the (\w+) field", t)
    fmt = re.search(r"displays the date as ([\w/.\-: ]+?)\.?$", t)
    for f in fields(d.body):
        if f.upper().startswith(m.group(1).upper()):
            return not fmt or fmt.group(1).strip() in f
    return False


# ====================================================================== vào ra


def grade(path, dang, de_bai, goc=None):
    """True / False; None nếu chưa có hàm chấm cho dạng này (học viên tự kiểm tra)."""
    fn = GRADERS.get(dang)
    if fn is None:
        return None
    path = Path(path)
    if dang == "encrypt":
        return path.read_bytes()[:4] == b"\xd0\xcf\x11\xe0"
    return fn(Doc(path), de_bai, goc or {})
