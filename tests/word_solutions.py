"""Lời giải mô phỏng cho các dạng câu Word của bộ đề nhập (dùng trong test).

Mỗi hàm `SOL[dang](pkg, text)` sửa XML trong file .docx giống cách Word lưu khi
học viên làm đúng câu đó. Không có Office trên máy test nên đây là mô phỏng.
"""
from __future__ import annotations

import copy
import re
import zipfile
from pathlib import Path

from lxml import etree

from mos import word_auto as A

NS = A.NS
W = A.W


def q(tag: str):
    pre, _, name = tag.partition(":")
    return "{%s}%s" % (NS[pre], name)


def el(tag, parent=None, **attrs):
    e = etree.Element(q(tag)) if parent is None else etree.SubElement(parent, q(tag))
    for k, v in attrs.items():
        pre, _, name = k.partition("_")
        e.set(q(f"{pre}:{name}") if name else k, str(v))
    return e


def wattr(e, name, val):
    e.set(W + name, str(val))
    return e


class Pkg:
    """Gói .docx có thể sửa: d = word_auto.Doc sống trên cùng cây XML."""

    def __init__(self, path: Path):
        self.path = Path(path)
        self.d = A.Doc(path)
        self.extra: dict[str, bytes] = {}

    def x(self, name, template=None):
        root = self.d.xml(name)
        if root is None and template is not None:
            self.d.files[name] = template.encode("utf-8")
            self.d._xml.pop(name, None)
            root = self.d.xml(name)
        return root

    def put(self, name, data: bytes):
        self.d.files[name] = data
        self.d._xml.pop(name, None)

    def save(self):
        out = {}
        for name, data in self.d.files.items():
            root = self.d._xml.get(name)
            out[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                                       standalone=True) if root is not None else data
        with zipfile.ZipFile(self.path, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in out.items():
                z.writestr(name, data)

    # ---------- tiện ích
    @property
    def body(self):
        return self.d.body

    def settings(self):
        return self.x("word/settings.xml")

    def rel(self, part, rtype, target, external=False):
        folder, _, name = part.rpartition("/")
        rels_name = f"{folder}/_rels/{name}.rels"
        root = self.x(rels_name, '<Relationships xmlns="%s"/>' % NS["pr"])
        ids = {r.get("Id") for r in root}
        n = 1
        while f"rId{n}" in ids:
            n += 1
        r = etree.SubElement(root, "{%s}Relationship" % NS["pr"])
        r.set("Id", f"rId{n}")
        r.set("Type", rtype)
        r.set("Target", target)
        if external:
            r.set("TargetMode", "External")
        return f"rId{n}"


RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"


def ppr(p):
    e = p.find(W + "pPr")
    if e is None:
        e = etree.Element(W + "pPr")
        p.insert(0, e)
    return e


PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
             "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "spacing", "ind", "jc",
             "outlineLvl", "rPr", "sectPr"]


def pset(p, name, **attrs):
    """Đặt phần tử con của pPr (đúng thứ tự schema)."""
    pr = ppr(p)
    e = pr.find(W + name)
    if e is None:
        e = etree.Element(W + name)
        idx = PPR_ORDER.index(name) if name in PPR_ORDER else len(PPR_ORDER)
        pos = 0
        for i, ch in enumerate(pr):
            ln = etree.QName(ch).localname
            if ln in PPR_ORDER and PPR_ORDER.index(ln) < idx:
                pos = i + 1
        pr.insert(pos, e)
    for k, v in attrs.items():
        wattr(e, k, v)
    return e


def rpr(r):
    e = r.find(W + "rPr")
    if e is None:
        e = etree.Element(W + "rPr")
        r.insert(0, e)
    return e


def rset(r, name, **attrs):
    e = rpr(r).find(W + name)
    if e is None:
        e = etree.SubElement(rpr(r), W + name)
    for k, v in attrs.items():
        wattr(e, k, v)
    return e


def run(text, parent=None):
    r = etree.Element(W + "r") if parent is None else etree.SubElement(parent, W + "r")
    t = etree.SubElement(r, W + "t")
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return r


def split_phrase(p, phrase):
    """Tách run để `phrase` nằm trọn trong các run riêng; trả về danh sách run đó."""
    runs = A.text_runs(p)
    full = "".join("".join(t.text or "" for t in r if t.tag == A.T) for r in runs)
    i = full.find(phrase)
    assert i >= 0, (phrase, full)
    j = i + len(phrase)
    pos, out = 0, []
    for r in runs:
        t = "".join(x.text or "" for x in r if x.tag == A.T)
        a, b = pos, pos + len(t)
        pos = b
        if b <= i or a >= j:
            continue
        s, e = max(a, i) - a, min(b, j) - a
        parts = [(t[:s], False), (t[s:e], True), (t[e:], False)]
        parent, idx = r.getparent(), r.getparent().index(r)
        for txt, mid in parts:
            if not txt:
                continue
            nr = copy.deepcopy(r)
            for x in list(nr):
                if x.tag != W + "rPr":
                    nr.remove(x)
            tt = etree.SubElement(nr, A.T)
            tt.text = txt
            tt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            parent.insert(idx, nr)
            idx += 1
            if mid:
                out.append(nr)
        parent.remove(r)
    return out


def para_with(pkg, text, start=False):
    p = A.find_para(pkg.d, text, start=start)
    if p is None:
        for x in A.all_paras(pkg.body):
            if A.norm(text.rstrip(" …").rstrip(".")) in A.norm(A.ptext(x)):
                return x
    assert p is not None, text
    return p


def ensure_style(pkg, sid, name, kind="paragraph"):
    styles = pkg.x("word/styles.xml")
    st = pkg.d.styles().get(sid)
    if st is None:
        st = el("w:style", styles)
        wattr(st, "type", kind)
        wattr(st, "styleId", sid)
        wattr(el("w:name", st), "val", name)
        pkg.d._styles = None
    return st


def secs(pkg):
    return A.sect_prs(pkg.d)


def sec_obj(pkg, text, kind):
    drs = A.sec_objects(pkg.d, text, kind)
    assert drs, (kind, text)
    return drs[0]


def quotes(t):
    return A.quotes(t)


# ====================================================================== lời giải

SOL = {}


def sol(*names):
    def deco(fn):
        for n in names:
            SOL[n] = fn
        return fn
    return deco


def _core(pkg, tag, value):
    root = pkg.x("docProps/core.xml")
    for e in root:
        if etree.QName(e).localname == tag:
            e.text = value
            return
    ns = {"category": "cp", "keywords": "cp", "contentStatus": "cp", "subject": "dc", "title": "dc",
          "creator": "dc"}[tag]
    e = etree.SubElement(root, "{%s}%s" % (NS[ns], tag))
    e.text = value


@sol("prop_category")
def _(pkg, t):
    _core(pkg, "category", quotes(t)[0])


@sol("prop_subject")
def _(pkg, t):
    _core(pkg, "subject", quotes(t)[0])


@sol("prop_tags")
def _(pkg, t):
    _core(pkg, "keywords", quotes(t)[0])


@sol("prop_title")
def _(pkg, t):
    _core(pkg, "title", quotes(t)[0])


@sol("prop_company")
def _(pkg, t):
    root = pkg.x("docProps/app.xml")
    comp = root.find("ep:Company", NS)
    if comp is None:
        comp = etree.SubElement(root, "{%s}Company" % NS["ep"])
    comp.text = re.search(r"to (.+?)\.$", t).group(1)


@sol("inspect_props")
def _(pkg, t):
    root = pkg.x("docProps/core.xml")
    for e in list(root):
        root.remove(e)
    app = pkg.x("docProps/app.xml")
    for e in app.findall("ep:Company", NS):
        app.remove(e)
    el("w:removePersonalInformation", pkg.settings())


@sol("inspect_hf")
def _(pkg, t):
    for sp in secs(pkg):
        for ref in sp.findall(W + "headerReference") + sp.findall(W + "footerReference"):
            sp.remove(ref)


@sol("access_table")
def _(pkg, t):
    for tb in pkg.body.iter(A.TBL):
        tr = tb.find(W + "tr")
        trpr = tr.find(W + "trPr")
        if trpr is None:
            trpr = etree.Element(W + "trPr")
            tr.insert(0, trpr)
        if trpr.find(W + "tblHeader") is None:
            el("w:tblHeader", trpr)


@sol("access_alt", "alt_text")
def _(pkg, t):
    kind_drs = A.sec_objects(pkg.d, t, "pic") if A.sec_name(t) else A.drawings([pkg.body], "pic")
    A.docpr(kind_drs[0]).set("descr", quotes(t)[-1])


@sol("chart_alt")
def _(pkg, t):
    A.docpr(sec_obj(pkg, t, "chart")).set("descr", quotes(t)[0])


@sol("sa_alt")
def _(pkg, t):
    A.docpr(sec_obj(pkg, t, "dgm")).set("descr", quotes(t)[0])


@sol("compat")
def _(pkg, t):
    for cs in pkg.settings().iter(W + "compatSetting"):
        if cs.get(W + "name") == "compatibilityMode":
            wattr(cs, "val", 15)


@sol("save_txt", "save_pdf", "save_template", "save_rtf", "save_97")
def _(pkg, t):
    ext = {"plain-text": ".txt", "PDF": ".pdf", "template": ".dotx", "Rich Text": ".rtf", "97-2003": ".doc"}
    e = next(v for k, v in ext.items() if k in t)
    data = b"\xd0\xcf\x11\xe0 fake doc" if e == ".doc" else b"copy"
    (pkg.path.parent / f"{quotes(t)[0]}{e}").write_bytes(data)


def _mar(pkg, top, lr):
    for sp in secs(pkg):
        m = sp.find(W + "pgMar")
        wattr(m, "top", int(top * 1440))
        wattr(m, "bottom", int(top * 1440))
        wattr(m, "left", int(lr * 1440))
        wattr(m, "right", int(lr * 1440))


@sol("margins")
def _(pkg, t):
    v = A.inches(t)
    _mar(pkg, v[0], v[1])


@sol("margins_preset")
def _(pkg, t):
    key = re.search(r"Apply the (\w+) margin", t, re.I).group(1).casefold()
    if key == "mirrored":
        _mar(pkg, 1, 1.25)
        el("w:mirrorMargins", pkg.settings())
    else:
        _mar(pkg, *A.MARGIN_PRESETS[key])


def _landscape(sp):
    sz = sp.find(W + "pgSz")
    w, h = sz.get(W + "w"), sz.get(W + "h")
    wattr(sz, "w", max(int(w), int(h)))
    wattr(sz, "h", min(int(w), int(h)))
    wattr(sz, "orient", "landscape")


@sol("orient_all")
def _(pkg, t):
    for sp in secs(pkg):
        _landscape(sp)


@sol("paper")
def _(pkg, t):
    w, h = A.PAPER[re.search(r"to ([\w ]+)\.", t).group(1).strip().casefold()]
    for sp in secs(pkg):
        sz = sp.find(W + "pgSz")
        wattr(sz, "w", w)
        wattr(sz, "h", h)


def _theme(pkg):
    return pkg.x(pkg.d.glob(r"word/theme/theme\d*\.xml")[0])


@sol("doc_theme")
def _(pkg, t):
    _theme(pkg).set("name", re.search(r"Apply the (.+?) theme", t, re.I).group(1))


@sol("theme_colors")
def _(pkg, t):
    m = re.search(r"to (.+?) and the theme fonts to (.+?)\.$", t)
    th = _theme(pkg)
    th.find(".//a:clrScheme", NS).set("name", m.group(1))
    fs = th.find(".//a:fontScheme", NS)
    fs.set("name", m.group(2))
    fs.find("a:majorFont/a:latin", NS).set("typeface", m.group(2))


@sol("custom_colors")
def _(pkg, t):
    m = re.search(r"uses (.+?) \(Standard Colors\) as the Accent (\d)", t)
    cs = _theme(pkg).find(".//a:clrScheme", NS)
    cs.set("name", quotes(t)[0])
    acc = cs.find(f"a:accent{m.group(2)}", NS)
    for ch in list(acc):
        acc.remove(ch)
    el("a:srgbClr", acc, val=A.color_spec(m.group(1))["hex"])


@sol("custom_fonts")
def _(pkg, t):
    m = re.search(r"uses (.+?) as the heading font and (.+?) as the body font", t)
    fs = _theme(pkg).find(".//a:fontScheme", NS)
    fs.set("name", quotes(t)[0])
    fs.find("a:majorFont/a:latin", NS).set("typeface", m.group(1))
    fs.find("a:minorFont/a:latin", NS).set("typeface", m.group(2))


def _normal_spacing(pkg, **attrs):
    styles = pkg.x("word/styles.xml")
    pr = styles.find("w:docDefaults/w:pPrDefault/w:pPr", NS)
    sp = pr.find(W + "spacing")
    if sp is None:
        sp = el("w:spacing", pr)
    for k, v in attrs.items():
        wattr(sp, k, v)
    normal = pkg.d.default_style()
    nsp = normal.find("w:pPr/w:spacing", NS)
    if nsp is not None:
        nsp.getparent().remove(nsp)


@sol("style_set")
def _(pkg, t):
    st = pkg.d.styles()["Heading1"]
    rp = st.find(W + "rPr")
    wattr(el("w:caps", rp), "val", 1)


@sol("para_spacing_doc")
def _(pkg, t):
    key = re.search(r"to (\w+)\.$", t).group(1).casefold()
    _normal_spacing(pkg, after=160, line=A.SPACING_PRESETS.get(key, 264), lineRule="auto")


def _hf_part(pkg, kind, xml_inner, typ="default"):
    n = 1
    while f"word/{kind}{n}.xml" in pkg.d.files:
        n += 1
    name = f"word/{kind}{n}.xml"
    tag = "hdr" if kind == "header" else "ftr"
    pkg.put(name, (f'<w:{tag} xmlns:w="{NS["w"]}" xmlns:v="{NS["v"]}" xmlns:r="{NS["r"]}">{xml_inner}'
                   f'</w:{tag}>').encode())
    rid = pkg.rel("word/document.xml", RT + kind, f"{kind}{n}.xml")
    for sp in secs(pkg):
        for ref in sp.findall(W + f"{kind}Reference"):
            if ref.get(W + "type") == typ:
                sp.remove(ref)
        ref = etree.Element(W + f"{kind}Reference")
        wattr(ref, "type", typ)
        ref.set(q("r:id"), rid)
        sp.insert(0, ref)
    return name


@sol("header_gallery", "footer_gallery")
def _(pkg, t):
    kind = "header" if " header" in t else "footer"
    _hf_part(pkg, kind, '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Headers"/></w:docPartObj>'
                        '</w:sdtPr><w:sdtContent><w:p><w:r><w:t>Banded</w:t></w:r></w:p></w:sdtContent></w:sdt>')
    for sp in secs(pkg):
        tp = sp.find(W + "titlePg")
        if "except page 1" in t and tp is None:
            sp.append(etree.Element(W + "titlePg"))
        elif "except page 1" not in t and tp is not None:
            sp.remove(tp)


@sol("page_numbers")
def _(pkg, t):
    _hf_part(pkg, "footer", '<w:p><w:fldSimple w:instr=" PAGE   \\* MERGEFORMAT "><w:r><w:t>1</w:t></w:r>'
                            '</w:fldSimple></w:p>')
    m = re.search(r"start at (\d+)", t)
    if m:
        for sp in secs(pkg):
            wattr(el("w:pgNumType", sp), "start", m.group(1))


@sol("watermark")
def _(pkg, t):
    txt = quotes(t)[0] if quotes(t) else re.sub(r"\s+\d+$", "", re.search(r"Add the (.+?) watermark", t).group(1))
    _hf_part(pkg, "header", f'<w:p><w:r><w:pict><v:shape id="PowerPlusWaterMarkObject1"><v:textpath '
                            f'style="font-family:&quot;Calibri&quot;" string="{txt}"/></v:shape></w:pict></w:r></w:p>')


@sol("page_color")
def _(pkg, t):
    spec = A.color_spec(re.search(r"to (.+?)\.$", t).group(1))
    bg = etree.Element(W + "background")
    wattr(bg, "color", "DEEAF6")
    wattr(bg, "themeColor", spec["theme"])
    wattr(bg, "themeTint", spec["tint"])
    pkg.d.doc.insert(0, bg)


@sol("page_border")
def _(pkg, t):
    m = re.search(r"Add a (.+?) pt (.+?) Box page border", t)
    spec, sz = A.color_spec(m.group(2)), A._border_sz(t)
    for sp in secs(pkg):
        pb = etree.Element(W + "pgBorders")
        if "first page only" in t:
            wattr(pb, "display", "firstPage")
        for side in ("top", "left", "bottom", "right"):
            b = el("w:" + side, pb)
            wattr(b, "val", "single")
            wattr(b, "sz", sz)
            if spec.get("auto"):
                wattr(b, "color", "auto")
            else:
                wattr(b, "color", "2F5496")
                wattr(b, "themeColor", spec["theme"])
                if spec.get("shade"):
                    wattr(b, "themeShade", spec["shade"])
        sp.insert(len(sp.findall(W + "headerReference") + sp.findall(W + "footerReference")) + 3, pb)


@sol("hyphenation")
def _(pkg, t):
    el("w:autoHyphenation", pkg.settings())


@sol("hyperlink_url")
def _(pkg, t):
    phrase = quotes(t)[-1]
    url = re.search(r"links to (https?://\S+?)\.?(?:\s|$)", t).group(1)
    p = [x for x in A.sec_paras(pkg.d, A.sec_name(t)) if phrase in A.ptext(x)][0]
    runs = split_phrase(p, phrase)
    rid = pkg.rel("word/document.xml", RT + "hyperlink", url, external=True)
    h = etree.Element(W + "hyperlink")
    h.set(q("r:id"), rid)
    runs[0].addprevious(h)
    for r in runs:
        h.append(r)


@sol("hyperlink_place")
def _(pkg, t):
    phrase, target = quotes(t)[1], quotes(t)[2]
    hp = A.heading_para(pkg.d, target)
    bs = el("w:bookmarkStart", None)
    wattr(bs, "id", 90)
    wattr(bs, "name", "_Places")
    hp.insert(1, bs)
    be = etree.SubElement(hp, W + "bookmarkEnd")
    wattr(be, "id", 90)
    p = [x for x in A.all_paras(pkg.body) if phrase in A.ptext(x)][0]
    runs = split_phrase(p, phrase)
    h = etree.Element(W + "hyperlink")
    wattr(h, "anchor", "_Places")
    runs[0].addprevious(h)
    for r in runs:
        h.append(r)


@sol("bookmark")
def _(pkg, t):
    p = para_with(pkg, quotes(t)[0], start=True)
    bs = etree.Element(W + "bookmarkStart")
    wattr(bs, "id", 91)
    wattr(bs, "name", quotes(t)[-1])
    be = etree.Element(W + "bookmarkEnd")
    wattr(be, "id", 91)
    idx = 1 if p.find(W + "pPr") is not None else 0
    p.insert(idx, be)
    p.insert(idx, bs)


@sol("mark_final")
def _(pkg, t):
    pkg.put("docProps/custom.xml", b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/'
            b'custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            b'<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="_MarkAsFinal"><vt:bool>true'
            b'</vt:bool></property></Properties>')
    _core(pkg, "contentStatus", "Final")


def _protect(pkg, **attrs):
    st = pkg.settings()
    pr = etree.Element(W + "documentProtection")
    for k, v in attrs.items():
        wattr(pr, k, v)
    st.insert(1, pr)


@sol("restrict_comments")
def _(pkg, t):
    _protect(pkg, edit="comments", enforcement=1)


@sol("restrict_format")
def _(pkg, t):
    _protect(pkg, formatting=1, enforcement=0)


@sol("lock_tracking", "track_lock")
def _(pkg, t):
    _protect(pkg, edit="trackedChanges", enforcement=1, cryptProviderType="rsaAES", hashValue="abc=")
    if "Turn on" in t:
        el("w:trackRevisions", pkg.settings())


# ---------- chữ


def _replace_text(pkg, old, new, whole_word=False):
    for tt in pkg.body.iter(A.T):
        if tt.text and old in tt.text:
            tt.text = re.sub(rf"\b{re.escape(old)}\b ?" if whole_word else re.escape(old), new, tt.text)


@sol("find_delete")
def _(pkg, t):
    _replace_text(pkg, quotes(t)[0], "", whole_word=True)


@sol("replace_all")
def _(pkg, t):
    _replace_text(pkg, quotes(t)[0], quotes(t)[1])


@sol("wildcard_replace")
def _(pkg, t):
    for old in quotes(t)[:-1]:
        _replace_text(pkg, old, quotes(t)[-1])


@sol("special_char")
def _(pkg, t):
    ch = re.search(r"\((.)\)", t).group(1)
    org = quotes(t)[-1]
    for tt in pkg.body.iter(A.T):
        if tt.text and org in tt.text:
            tt.text = tt.text.replace(org, org + ch, 1)
            return


@sol("symbol")
def _(pkg, t):
    m = re.search(r"Use the (.+?) font and character code [“\"](\w+)[”\"]", t)
    r = etree.Element(W + "r")
    s = el("w:sym", r)
    wattr(s, "font", m.group(1))
    wattr(s, "char", f"F0{int(m.group(2)):02X}")
    if "document title" in t:
        A.title_para(pkg.d).append(r)
    else:
        p = para_with(pkg, quotes(t)[1])
        split_phrase(p, quotes(t)[1])[0].addprevious(r)


@sol("format_painter")
def _(pkg, t):
    ps = A.sec_paras(pkg.d, A.sec_name(t))
    src = A.text_runs(ps[0])[0].find(W + "rPr")
    for r in A.text_runs(ps[1]):
        old = r.find(W + "rPr")
        if old is not None:
            r.remove(old)
        r.insert(0, copy.deepcopy(src))


@sol("line_spacing_doc")
def _(pkg, t):
    v = float(re.search(r"to ([\d.]+) lines", t).group(1))
    for p in A.all_paras(pkg.body):
        pset(p, "spacing", line=int(round(v * 240)), lineRule="auto")


@sol("line_spacing_exact")
def _(pkg, t):
    v = int(re.search(r"exactly (\d+) pt", t).group(1)) * 20
    for p in A.sec_paras(pkg.d, A.sec_name(t))[:2]:
        pset(p, "spacing", line=v, lineRule="exact")


@sol("para_spacing")
def _(pkg, t):
    m = re.search(r"to (\d+) pt and the spacing after to (\d+) pt", t)
    for p in A.sec_paras(pkg.d, A.sec_name(t))[:2]:
        pset(p, "spacing", before=int(m.group(1)) * 20, after=int(m.group(2)) * 20)


@sol("indent")
def _(pkg, t):
    v = int(round(A.inches(t)[0] * 1440))
    for p in A.sec_paras(pkg.d, A.sec_name(t))[:2]:
        pset(p, "ind", **({"firstLine": v} if "first line" in t else {"left": v}))


@sol("clear_format")
def _(pkg, t):
    p = para_with(pkg, quotes(t)[0], start=True)
    pr = p.find(W + "pPr")
    if pr is not None:
        p.remove(pr)
    for r in A.text_runs(p):
        x = r.find(W + "rPr")
        if x is not None:
            r.remove(x)


def _style_id(name):
    return re.sub(r"[^A-Za-z0-9]", "", name)


@sol("para_style")
def _(pkg, t):
    name = re.search(r"Apply the (.+?) style", t, re.I).group(1)
    ensure_style(pkg, _style_id(name), name)
    pset(para_with(pkg, quotes(t)[0], start=True), "pStyle", val=_style_id(name))


@sol("heading_style")
def _(pkg, t):
    name = re.search(r"Apply the (.+?) style", t, re.I).group(1)
    ensure_style(pkg, _style_id(name), name.lower() if name.startswith("Heading") else name)
    pset(A.heading_para(pkg.d, quotes(t)[0]), "pStyle", val=_style_id(name))


def _apply_char_style(pkg, sentence, sid):
    p = para_with(pkg, sentence.rstrip("."))
    for r in split_phrase(p, sentence):
        rs = rpr(r).find(W + "rStyle")
        if rs is None:
            rs = etree.Element(W + "rStyle")
            rpr(r).insert(0, rs)
        wattr(rs, "val", sid)


@sol("char_style")
def _(pkg, t):
    name = re.search(r"apply the (.+?) style to it", t, re.I).group(1)
    ensure_style(pkg, _style_id(name), name, "character")
    _apply_char_style(pkg, quotes(t)[0], _style_id(name))


@sol("char_style_new")
def _(pkg, t):
    name, sentence = quotes(t)[0], quotes(t)[1]
    color = A.color_spec(re.search(r"uses bold, (.+?) font color", t).group(1))["hex"]
    st = ensure_style(pkg, _style_id(name), name, "character")
    rp = el("w:rPr", st)
    el("w:b", rp)
    wattr(el("w:color", rp), "val", color)
    _apply_char_style(pkg, sentence, _style_id(name))


@sol("para_style_new")
def _(pkg, t):
    name, start = quotes(t)[0], quotes(t)[1]
    m = re.search(r"uses the (.+?) font, (\d+) pt, italic, with (\d+) pt of spacing after", t)
    st = ensure_style(pkg, "InfoBox", name)
    pp = el("w:pPr", st)
    wattr(el("w:spacing", pp), "after", int(m.group(3)) * 20)
    rp = el("w:rPr", st)
    f = el("w:rFonts", rp)
    wattr(f, "ascii", m.group(1))
    wattr(f, "hAnsi", m.group(1))
    el("w:i", rp)
    wattr(el("w:sz", rp), "val", int(m.group(2)) * 2)
    pset(para_with(pkg, start, start=True), "pStyle", val="InfoBox")


@sol("modify_style")
def _(pkg, t):
    m = re.search(r"Modify the (.+?) style so that it uses (\d+) pt, (bold, )?(.+?) font color", t)
    st = pkg.d.style_by_name(m.group(1))
    rp = st.find(W + "rPr")
    for x in list(rp):
        if etree.QName(x).localname in ("sz", "b", "color"):
            rp.remove(x)
    el("w:b", rp)
    spec = A.color_spec(m.group(4))
    c = el("w:color", rp)
    wattr(c, "val", "C45911")
    wattr(c, "themeColor", spec["theme"])
    if spec.get("shade"):
        wattr(c, "themeShade", spec["shade"])
    wattr(el("w:sz", rp), "val", int(m.group(2)) * 2)


@sol("replace_style")
def _(pkg, t):
    old, new = re.findall(r"(Heading \d)", t)
    ensure_style(pkg, _style_id(new), new.lower())
    for p in A.all_paras(pkg.body):
        if A.pstyle(pkg.d, p) == A.norm(old):
            pset(p, "pStyle", val=_style_id(new))


@sol("change_case")
def _(pkg, t):
    head = quotes(t)[0]
    mode = t.rsplit(" to ", 1)[-1].rstrip(".")
    new = head.upper() if mode == "UPPERCASE" else " ".join(w[:1].upper() + w[1:].lower() for w in head.split())
    p = A.heading_para(pkg.d, head)
    runs = A.text_runs(p)
    runs[0].find(A.T).text = new
    for r in runs[1:]:
        p.remove(r)


@sol("font_format")
def _(pkg, t):
    m = re.search(r"to (.+?), (\d+) pt", t)
    for r in A.text_runs(para_with(pkg, quotes(t)[0])):
        f = rset(r, "rFonts", ascii=m.group(1), hAnsi=m.group(1))
        f.getparent().remove(f)
        rpr(r).insert(0, f)
        rset(r, "sz", val=int(m.group(2)) * 2)


@sol("highlight")
def _(pkg, t):
    s = quotes(t)[-1]
    for r in split_phrase(para_with(pkg, s.rstrip(".")), s):
        rset(r, "highlight", val="green")


@sol("para_shading")
def _(pkg, t):
    spec = A.color_spec(re.search(r"Apply (.+?) shading", t).group(1))
    pset(para_with(pkg, quotes(t)[0], start=True), "shd", val="clear", color="auto", fill="FFF2CC",
         themeFill=spec["theme"], themeFillTint=spec["tint"])


@sol("text_effect")
def _(pkg, t):
    desc = re.search(r"Apply the (.+?) text effect", t, re.I).group(1).casefold()
    p = A.target_para(pkg.d, t)
    for r in A.text_runs(p):
        rp = rpr(r)
        if "shadow" in desc:
            el("w14:shadow", rp, w14_blurRad=38100)
        if "outline" in desc:
            el("w14:textOutline", rp)
        if "glow" in desc:
            el("w14:glow", rp, w14_rad=63500)
        el("w14:textFill", rp)


@sol("keep_next")
def _(pkg, t):
    p = A.heading_para(pkg.d, quotes(t)[0])
    pset(p, "keepNext")
    pset(p, "keepLines")


@sol("widow")
def _(pkg, t):
    for p in A.sec_paras(pkg.d, A.sec_name(t))[:2]:
        pset(p, "keepLines")
        pset(p, "widowControl", val=0)


@sol("page_break")
def _(pkg, t):
    h = A.heading_para(pkg.d, quotes(t)[0])
    p = etree.Element(W + "p")
    wattr(el("w:br", el("w:r", p)), "type", "page")
    h.addprevious(p)


def _sect_after(p, typ, pkg, **pg):
    sp = copy.deepcopy(A.final_sect(pkg.d))
    for x in sp.findall(W + "type"):
        sp.remove(x)
    tp = etree.Element(W + "type")
    wattr(tp, "val", typ)
    sp.insert(len(sp.findall(W + "headerReference") + sp.findall(W + "footerReference")) +
              len(sp.findall(W + "footnotePr")), tp)
    pset(p, "sectPr").getparent().replace(ppr(p).find(W + "sectPr"), sp)
    return sp


@sol("sbreak_cont")
def _(pkg, t):
    _sect_after(para_with(pkg, quotes(t)[0].rstrip(".")), "continuous", pkg)


@sol("sbreak_next")
def _(pkg, t):
    h = A.heading_para(pkg.d, quotes(t)[0])
    p = etree.Element(W + "p")
    h.addprevious(p)
    _sect_after(p, "nextPage", pkg)


@sol("orient_sec")
def _(pkg, t):
    h = A.heading_para(pkg.d, quotes(t)[0])
    before = etree.Element(W + "p")
    h.addprevious(before)
    _sect_after(before, "nextPage", pkg)
    after = etree.Element(W + "p")
    h.addnext(after)
    sp = _sect_after(after, "nextPage", pkg)
    _landscape(sp)


@sol("columns")
def _(pkg, t):
    ps = A.sec_paras(pkg.d, A.sec_name(t))[:2]
    n = re.search(r"into (\d) columns", t).group(1)
    before = etree.Element(W + "p")
    ps[0].addprevious(before)
    _sect_after(before, "continuous", pkg)
    sp = _sect_after(ps[1], "continuous", pkg)
    cols = sp.find(W + "cols")
    if cols is None:
        cols = el("w:cols", sp)
    wattr(cols, "num", n)
    if "line between" in t:
        wattr(cols, "sep", 1)
    if A.inches(t):
        wattr(cols, "space", int(round(A.inches(t)[0] * 1440)))


@sol("dropcap")
def _(pkg, t):
    p = A.sec_paras(pkg.d, A.sec_name(t))[0]
    cap = etree.Element(W + "p")
    pr = el("w:pPr", cap)
    fp = el("w:framePr", pr)
    wattr(fp, "dropCap", "drop")
    wattr(fp, "lines", re.search(r"drop (\d) lines", t).group(1))
    run("X", cap)
    p.addprevious(cap)


@sol("language")
def _(pkg, t):
    styles = pkg.x("word/styles.xml")
    rp = styles.find("w:docDefaults/w:rPrDefault/w:rPr", NS)
    lang = rp.find(W + "lang")
    if lang is None:
        lang = el("w:lang", rp)
    wattr(lang, "val", "en-GB")
    for r in pkg.body.iter(W + "lang"):
        wattr(r, "val", "en-GB")


@sol("line_numbers")
def _(pkg, t):
    for sp in secs(pkg):
        ln = etree.Element(W + "lnNumType")
        wattr(ln, "countBy", 1)
        sp.insert(len(sp) - 1, ln)


@sol("replace_color")
def _(pkg, t):
    new = A.color_spec(re.search(r"change the font color to (.+?)\.$", t).group(1))["hex"]
    for c in pkg.body.iter(W + "color"):
        if (c.get(W + "val") or "").upper() == "C00000":
            wattr(c, "val", new)


# ---------- bảng


def _tbl(pkg, t):
    tb = A._table(pkg.d, t)
    assert tb is not None, t
    return tb


def _set_text(tc, text):
    ps = tc.findall(W + "p")
    for p in ps[1:]:
        tc.remove(p)
    p = ps[0]
    for x in list(p):
        if x.tag != W + "pPr":
            p.remove(x)
    if text:
        run(text, p)


@sol("add_row")
def _(pkg, t):
    rows = A._rows(_tbl(pkg, t))
    new = copy.deepcopy(rows[-1])
    for tc, v in zip(A._cells(new), quotes(t)[1:]):
        _set_text(tc, v)
    rows[-1].addnext(new)


@sol("table_insert_col")
def _(pkg, t):
    after, new = quotes(t)[1], quotes(t)[2]
    tb = _tbl(pkg, t)
    idx = [A.norm(A._ctext(c)) for c in A._cells(A._rows(tb)[0])].index(A.norm(after))
    for i, tr in enumerate(A._rows(tb)):
        c = copy.deepcopy(A._cells(tr)[idx])
        _set_text(c, new if i == 0 else "")
        A._cells(tr)[idx].addnext(c)


@sol("delete_col")
def _(pkg, t):
    tb = _tbl(pkg, t)
    idx = [A.norm(A._ctext(c)) for c in A._cells(A._rows(tb)[0])].index(A.norm(quotes(t)[1]))
    for tr in A._rows(tb):
        c = A._cells(tr)[idx]
        tr.remove(c)
    grid = tb.find(W + "tblGrid")
    grid.remove(grid.findall(W + "gridCol")[idx])


@sol("merge_cells")
def _(pkg, t):
    tb = _tbl(pkg, t)
    tr = A._rows(tb)[0]
    cells = A._cells(tr)
    for c in cells[1:]:
        tr.remove(c)
    pr = cells[0].find(W + "tcPr")
    wattr(el("w:gridSpan", pr), "val", len(tb.findall(f"{W}tblGrid/{W}gridCol")))


@sol("split_cell")
def _(pkg, t):
    tb = _tbl(pkg, t)
    for c in A._cells(A._rows(tb)[0]):
        if A.norm(A._ctext(c)) == A.norm(quotes(t)[1]):
            n = copy.deepcopy(c)
            _set_text(n, "")
            c.addnext(n)
            return
    raise AssertionError("không thấy ô")


@sol("split_table")
def _(pkg, t):
    tb = _tbl(pkg, t)
    rows = A._rows(tb)
    i = next(k for k, r in enumerate(rows) if A.norm(A._ctext(A._cells(r)[0])).startswith(A.norm(quotes(t)[1])))
    new = copy.deepcopy(tb)
    for r in A._rows(new)[:i]:
        new.remove(r)
    for r in rows[i:]:
        tb.remove(r)
    tb.addnext(etree.Element(W + "p"))
    tb.getnext().addnext(new)


@sol("table_to_text", "info_to_text")
def _(pkg, t):
    tb = _tbl(pkg, t)
    sep = "," if "commas" in t else "\t"
    for r in A._rows(tb):
        p = etree.Element(W + "p")
        for i, c in enumerate(A._cells(r)):
            if i:
                if sep == "\t":
                    etree.SubElement(etree.SubElement(p, W + "r"), W + "tab")
                else:
                    run(",", p)
            run(A._ctext(c), p)
        tb.addprevious(p)
    tb.getparent().remove(tb)


@sol("text_to_table")
def _(pkg, t):
    ps = [p for p in A.sec_paras(pkg.d, A.sec_name(t)) if "\t" in A.ptext(p)]
    tb = etree.Element(A.TBL)
    grid = el("w:tblGrid", tb)
    for _ in range(2):
        wattr(el("w:gridCol", grid), "w", 4680)
    for p in ps:
        tr = el("w:tr", tb)
        for v in A.ptext(p).split("\t"):
            tc = el("w:tc", tr)
            run(v, el("w:p", tc))
    ps[0].addprevious(tb)
    for p in ps:
        p.getparent().remove(p)


@sol("insert_table")
def _(pkg, t):
    m = re.search(r"has (\d+) columns and (\d+) rows", t)
    ncol, nrow = int(m.group(1)), int(m.group(2))
    tb = etree.Element(A.TBL)
    pr = el("w:tblPr", tb)
    tw = el("w:tblW", pr)
    wattr(tw, "w", 0)
    wattr(tw, "type", "auto")
    for i in range(nrow):
        tr = el("w:tr", tb)
        for j in range(ncol):
            p = el("w:p", el("w:tc", tr))
            if i == 0:
                run(quotes(t)[1 + j], p)
    A.heading_para(pkg.d, quotes(t)[0]).addnext(tb)


@sol("autofit")
def _(pkg, t):
    tb = _tbl(pkg, t)
    tw = tb.find(f"{W}tblPr/{W}tblW")
    if "Window" in t:
        wattr(tw, "w", 5000)
        wattr(tw, "type", "pct")
    else:
        wattr(tw, "w", 0)
        wattr(tw, "type", "auto")
        for c in tb.iter(W + "tcW"):
            wattr(c, "type", "auto")
            wattr(c, "w", 0)


@sol("col_width")
def _(pkg, t):
    v = int(round(A.inches(t)[0] * 1440))
    tb = _tbl(pkg, t)
    for c in tb.iter(W + "gridCol", W + "tcW"):
        wattr(c, "w", v)


@sol("table_rowheight")
def _(pkg, t):
    v = int(round(A.inches(t)[0] * 1440))
    for tr in A._rows(_tbl(pkg, t)):
        trpr = tr.find(W + "trPr")
        if trpr is None:
            trpr = etree.Element(W + "trPr")
            tr.insert(0, trpr)
        wattr(el("w:trHeight", trpr), "val", v)
        for c in A._cells(tr):
            wattr(el("w:vAlign", c.find(W + "tcPr")), "val", "center")


@sol("cell_align")
def _(pkg, t):
    m = re.search(r"to Align (Top|Center|Bottom) (Left|Center|Right)", t)
    for c in A._cells(A._rows(_tbl(pkg, t))[0]):
        if m.group(1) != "Top":
            wattr(el("w:vAlign", c.find(W + "tcPr")), "val", m.group(1).lower())
        for p in c.iter(W + "p"):
            pset(p, "jc", val=m.group(2).lower())


@sol("cell_spacing")
def _(pkg, t):
    pr = _tbl(pkg, t).find(W + "tblPr")
    cs = etree.Element(W + "tblCellSpacing")
    wattr(cs, "w", int(round(A.inches(t)[0] * 1440 / 2)))
    wattr(cs, "type", "dxa")
    pr.insert(1, cs)


@sol("repeat_header")
def _(pkg, t):
    tr = A._rows(_tbl(pkg, t))[0]
    trpr = tr.find(W + "trPr")
    if trpr is None:
        trpr = etree.Element(W + "trPr")
        tr.insert(0, trpr)
    if trpr.find(W + "tblHeader") is None:
        el("w:tblHeader", trpr)


@sol("sort_table")
def _(pkg, t):
    tb = _tbl(pkg, t)
    rows = A._rows(tb)
    head = [A.norm(A._ctext(c)) for c in A._cells(rows[0])]
    keys = [(head.index(A.norm(m.group(1))), "desc" in (m.group(2) or "").casefold())
            for m in re.finditer(r"by [“\"]([^”\"]+)[”\"] ?\(?(Ascending|Descending|in ascending order|"
                                 r"in descending order)?", t, re.I)]
    body = [r for r in rows[1:] if not A.norm(A._ctext(A._cells(r)[0])).startswith("total")]

    def conv(x):
        try:
            return (0, float(x.replace(",", "")))
        except ValueError:
            return (1, x.casefold())
    for idx, desc in reversed(keys):
        body.sort(key=lambda r: conv(A._ctext(A._cells(r)[idx])), reverse=desc)
    for r in body:
        tb.remove(r)
    anchor = rows[0]
    for r in body:
        anchor.addnext(r)
        anchor = r


@sol("table_formula")
def _(pkg, t):
    for r in A._rows(_tbl(pkg, t)):
        cells = A._cells(r)
        if A.norm(A._ctext(cells[0])).startswith("total"):
            p = cells[-1].find(W + "p")
            fs = el("w:fldSimple", p)
            wattr(fs, "instr", ' =SUM(ABOVE) \\# "#,##0.00" ')
            run("10.00", fs)


@sol("table_style", "info_style")
def _(pkg, t):
    name = re.search(r"Apply the (.+?) table style", t, re.I).group(1)
    sid = _style_id(name)
    ensure_style(pkg, sid, name, "table")
    tb = _tbl(pkg, t)
    pr = tb.find(W + "tblPr")
    st = pr.find(W + "tblStyle")
    if st is None:
        st = etree.Element(W + "tblStyle")
        pr.insert(0, st)
    wattr(st, "val", sid)
    lk = pr.find(W + "tblLook")
    wattr(lk, "firstColumn", 0 if "first column" in t else 1)
    wattr(lk, "firstRow", 0 if "Header Row" in t else 1)
    wattr(lk, "val", "0400" if "first column" in t else "0080")


@sol("info_noborder")
def _(pkg, t):
    tb = _tbl(pkg, t)
    for bd in tb.iter(W + "tblBorders", W + "tcBorders"):
        for x in bd:
            wattr(x, "val", "nil")
    pr = tb.find(W + "tblPr")
    if pr.find(W + "tblBorders") is None:
        bd = el("w:tblBorders", pr)
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            wattr(el("w:" + side, bd), "val", "nil")


@sol("table_borders")
def _(pkg, t):
    m = re.search(r"apply a (.+?) pt (.+?) outside border", t)
    spec, sz = A.color_spec(m.group(2)), A._border_sz(t)
    pr = _tbl(pkg, t).find(W + "tblPr")
    bd = pr.find(W + "tblBorders")
    if bd is None:
        bd = el("w:tblBorders", pr)
    for side in ("top", "left", "bottom", "right"):
        b = bd.find(W + side)
        if b is None:
            b = el("w:" + side, bd)
        wattr(b, "val", "single")
        wattr(b, "sz", sz)
        wattr(b, "color", "4472C4")
        wattr(b, "themeColor", spec["theme"])


@sol("table_shading")
def _(pkg, t):
    spec = A.color_spec(re.search(r"header row of the table to (.+?)\.$", t).group(1))
    for c in A._cells(A._rows(_tbl(pkg, t))[0]):
        pr = c.find(W + "tcPr")
        for s in pr.findall(W + "shd"):
            pr.remove(s)
        s = el("w:shd", pr)
        wattr(s, "val", "clear")
        wattr(s, "fill", "70AD47")
        wattr(s, "themeFill", spec["theme"])


@sol("alt_table")
def _(pkg, t):
    pr = _tbl(pkg, t).find(W + "tblPr")
    wattr(el("w:tblCaption", pr), "val", quotes(t)[0])


# ---------- danh sách


def _new_num(pkg, fmt, text, font=None, start=1):
    root = pkg.x("word/numbering.xml", f'<w:numbering xmlns:w="{NS["w"]}"/>')
    ids = [int(a.get(W + "abstractNumId")) for a in root.findall("w:abstractNum", NS)]
    aid = max(ids + [0]) + 1
    absn = etree.Element(W + "abstractNum")
    wattr(absn, "abstractNumId", aid)
    for lvl_i in range(3):
        lvl = el("w:lvl", absn)
        wattr(lvl, "ilvl", lvl_i)
        wattr(el("w:start", lvl), "val", start)
        wattr(el("w:numFmt", lvl), "val", fmt)
        wattr(el("w:lvlText", lvl), "val", text.replace("%1", f"%{lvl_i + 1}"))
        if font:
            f = el("w:rFonts", el("w:rPr", lvl))
            wattr(f, "ascii", font)
            wattr(f, "hAnsi", font)
    first_num = root.find("w:num", NS)
    if first_num is not None:
        first_num.addprevious(absn)
    else:
        root.append(absn)
    nids = [int(n.get(W + "numId")) for n in root.findall("w:num", NS)]
    nid = max(nids + [0]) + 1
    num = el("w:num", root)
    wattr(num, "numId", nid)
    wattr(el("w:abstractNumId", num), "val", aid)
    return nid


def _apply_num(p, nid, ilvl=0):
    npr = pset(p, "numPr")
    for x in list(npr):
        npr.remove(x)
    wattr(el("w:ilvl", npr), "val", ilvl)
    wattr(el("w:numId", npr), "val", nid)


@sol("bullets_from_paras")
def _(pkg, t):
    n = int(re.search(r"the (\d+) paragraphs", t).group(1))
    nid = _new_num(pkg, "bullet", "", "Symbol")
    for p in A.paras_from(pkg.d, A.sec_name(t), quotes(t)[1], n):
        _apply_num(p, nid)


@sol("numbers_from_paras")
def _(pkg, t):
    n = int(re.search(r"the (\d+) paragraphs", t).group(1))
    kind, text = A.fmt_spec(re.search(r"uses the (.+?) number format", t).group(1))
    nid = _new_num(pkg, kind, text)
    for p in A.paras_from(pkg.d, A.sec_name(t), quotes(t)[1], n):
        _apply_num(p, nid)


def _sec_list(pkg, t, numbered=True):
    out = []
    for p in A._list_paras(pkg.d, A.sec_name(t)):
        info = A.numbering(pkg.d, p)
        if (info["numFmt"] != "bullet") == numbered:
            out.append((p, info))
    return out


@sol("num_format")
def _(pkg, t):
    raw = re.search(r"numbered list to (.+?)\.?$", t).group(1)
    if raw.endswith(".."):
        raw = raw[:-1]
    kind, text = A.fmt_spec(raw)
    nid = _new_num(pkg, kind, text)
    for p, info in _sec_list(pkg, t):
        _apply_num(p, nid, info["ilvl"])


@sol("custom_bullet")
def _(pkg, t):
    m = re.search(r"from the (.+?) font and character code [“\"](\w+)[”\"]", t)
    code = int(m.group(2), 16 if "Emoji" in m.group(1) else 10)
    ch = chr(0xF000 + code) if code < 0x100 else chr(code)
    nid = _new_num(pkg, "bullet", ch, m.group(1))
    for p, info in _sec_list(pkg, t, numbered=False):
        _apply_num(p, nid, info["ilvl"])


@sol("list_level")
def _(pkg, t):
    lvl = int(re.search(r"to Level (\d)", t).group(1)) - 1
    for p in A.section(pkg.d, A.sec_name(t)):
        if p.tag == A.P and A.norm(A.ptext(p)) == A.norm(quotes(t)[1]):
            wattr(p.find(f"{W}pPr/{W}numPr/{W}ilvl"), "val", lvl)


@sol("restart_num")
def _(pkg, t):
    start = re.search(r"starts at (\d+)", t).group(1)
    items = [(p, i) for p, i in _sec_list(pkg, t) if i["ilvl"] == 0]
    root = pkg.x("word/numbering.xml")
    nids = [int(n.get(W + "numId")) for n in root.findall("w:num", NS)]
    nid = max(nids) + 1
    num = el("w:num", root)
    wattr(num, "numId", nid)
    wattr(el("w:abstractNumId", num), "val", items[0][1]["abs"])
    ov = el("w:lvlOverride", num)
    wattr(ov, "ilvl", 0)
    wattr(el("w:startOverride", ov), "val", start)
    for p, info in _sec_list(pkg, t):
        wattr(p.find(f"{W}pPr/{W}numPr/{W}numId"), "val", nid)


@sol("continue_num")
def _(pkg, t):
    items = _sec_list(pkg, t)
    first = items[0][1]["numId"]
    for p, info in items:
        wattr(p.find(f"{W}pPr/{W}numPr/{W}numId"), "val", first)


# ---------- tham chiếu


def _field(p, instr, result="x"):
    fs = el("w:fldSimple", p)
    wattr(fs, "instr", f" {instr} ")
    run(result, fs)
    return fs


@sol("footnote")
def _(pkg, t):
    kind = "endnote" if "endnote" in t.split(".")[0] else "footnote"
    h = A.heading_para(pkg.d, quotes(t)[0])
    r = el("w:r", h)
    ref = el(f"w:{kind}Reference", r)
    wattr(ref, "id", 5)
    root = pkg.x(f"word/{kind}s.xml", f'<w:{kind}s xmlns:w="{NS["w"]}"/>')
    note = el(f"w:{kind}", root)
    wattr(note, "id", 5)
    run(quotes(t)[-1], el("w:p", note))


@sol("convert_notes")
def _(pkg, t):
    src, dst = ("footnote", "endnote") if "footnotes to endnotes" in t else ("endnote", "footnote")
    for ref in list(pkg.body.iter(W + f"{src}Reference")):
        ref.tag = W + f"{dst}Reference"


@sol("fn_format")
def _(pkg, t):
    sp = A.final_sect(pkg.d)
    pr = etree.Element(W + "footnotePr")
    wattr(el("w:numFmt", pr), "val", "lowerRoman")
    sp.insert(0, pr)


def _blank_below_banner(pkg):
    for b in A.blocks(pkg.d):
        if b.tag == A.P and not A.ptext(b).strip():
            return b
    raise AssertionError


@sol("toc_insert")
def _(pkg, t):
    _field(_blank_below_banner(pkg), 'TOC \\o "1-3" \\h \\z \\u')


@sol("custom_toc")
def _(pkg, t):
    _field(_blank_below_banner(pkg), 'TOC \\o "1-3" \\h \\z \\t "Subtitle,4"')


@sol("toc_modify")
def _(pkg, t):
    n = "2" if "Heading 2" in t else "1"
    for fs in pkg.body.iter(W + "instrText"):
        if fs.text and "TOC" in fs.text:
            fs.text = re.sub(r'\\o "[^"]*"', f'\\\\o "1-{n}"', fs.text)
    for fs in pkg.body.iter(W + "fldSimple"):
        if "TOC" in fs.get(W + "instr", ""):
            fs.set(W + "instr", re.sub(r'\\o "[^"]*"', f'\\\\o "1-{n}"', fs.get(W + "instr")))


@sol("table_of_figures")
def _(pkg, t):
    p = A.heading_para(pkg.d, "List of Tables").getnext()
    _field(p, 'TOC \\h \\z \\c "Table"')


@sol("caption_pic", "caption_table")
def _(pkg, t):
    label = "Table" if "Table label" in t else "Figure"
    p = etree.Element(W + "p")
    run(f"{label} ", p)
    _field(p, f"SEQ {label} \\* ARABIC", "1")
    run(quotes(t)[-1][len("Figure 1"):] if label == "Figure" else quotes(t)[-1], p)
    if label == "Table":
        tb = A.sec_tables(pkg.d, A.sec_name(t))[0]
        (tb.addprevious if "above" in t else tb.addnext)(p)
    else:
        dr = sec_obj(pkg, t, "pic")
        next(dr.iterancestors(A.P)).addnext(p)


@sol("cross_ref")
def _(pkg, t):
    p = [x for x in A.sec_paras(pkg.d, A.sec_name(t)) if "see" in A.ptext(x)][0]
    _field(p, "REF _Ref123456 \\h", "Table 1")


@sol("citation")
def _(pkg, t):
    idx = 0 if "first paragraph" in t else 1
    p = A.sec_paras(pkg.d, A.sec_name(t))[idx]
    _field(p, f"CITATION {quotes(t)[-1]} \\l 1033", f"({quotes(t)[-1]})")


@sol("citation_source")
def _(pkg, t):
    title = re.search(r"Title: (.+?);", t).group(1)
    p = A.sec_paras(pkg.d, A.sec_name(t))[1]
    _field(p, "CITATION Min22 \\l 1033", "(Nguyen, 2022)")
    name = next(n for n in pkg.d.glob(r"customXml/item\d+\.xml") if "Sources" in pkg.d.raw(n))
    pkg.put(name, pkg.d.raw(name).replace("/>", f'><b:Source><b:Title>{title}</b:Title></b:Source></b:Sources>', 1)
            .encode())


@sol("biblio_style")
def _(pkg, t):
    style = re.search(r"to (\w+)\.$", t).group(1)
    name = next(n for n in pkg.d.glob(r"customXml/item\d+\.xml") if "Sources" in pkg.d.raw(n))
    raw = re.sub(r'StyleName="[^"]*"', f'StyleName="{style}"', pkg.d.raw(name))
    pkg.put(name, re.sub(r'SelectedStyle="[^"]*"', f'SelectedStyle="\\\\{style}.XSL"', raw).encode())


@sol("index")
def _(pkg, t):
    phrase = quotes(t)[0]
    p = [x for x in A.all_paras(pkg.body) if phrase in A.ptext(x)][0]
    r = el("w:r", p)
    etree.SubElement(r, W + "fldChar").set(W + "fldCharType", "begin")
    it = etree.SubElement(el("w:r", p), W + "instrText")
    it.text = f' XE "{phrase}" '
    etree.SubElement(el("w:r", p), W + "fldChar").set(W + "fldCharType", "end")
    blank = [b for b in A.blocks(pkg.d) if b.tag == A.P and not A.ptext(b).strip()][-1]
    _field(blank, 'INDEX \\e "\t" \\c "2" \\z "1033"')


@sol("fields")
def _(pkg, t):
    name = re.search(r"insert the (\w+) field", t).group(1)
    fmt = re.search(r"displays the date as ([\w/.\-: ]+?)\.?$", t)
    blank = [b for b in A.blocks(pkg.d) if b.tag == A.P][-1]
    _field(blank, f'{name.upper()}' + (f' \\@ "{fmt.group(1)}"' if fmt else "") + " \\* MERGEFORMAT")


# ---------- đồ họa


def _set_extent(dr, cx=None, cy=None):
    e = dr.find("wp:extent", NS)
    for x in [e] + list(dr.iter(q("a:ext"))):
        if x.getparent().tag == q("a:xfrm") or x is e:
            if cx:
                x.set("cx", str(int(cx)))
            if cy:
                x.set("cy", str(int(cy)))


@sol("pic_size")
def _(pkg, t):
    dr = sec_obj(pkg, t, "pic")
    cx, cy = A.extent(dr)
    v = A.inches(t)[0] * 914400
    if "width" in t:
        _set_extent(dr, v, cy * v / cx)
    else:
        _set_extent(dr, cx * v / cy, v)


@sol("chart_size", "sa_size")
def _(pkg, t):
    kind = "dgm" if "SmartArt" in t else "chart"
    h = float(re.search(r"(\d+(?:\.\d+)?)\" \([^)]*\) high", t).group(1))
    w = float(re.search(r"(\d+(?:\.\d+)?)\" \([^)]*\) wide", t).group(1))
    _set_extent(sec_obj(pkg, t, kind), w * 914400, h * 914400)


def _sppr(pkg, t):
    return A._sppr(sec_obj(pkg, t, "pic"))


def _add_effect(sp, desc):
    low = desc.casefold()
    lst = sp.find("a:effectLst", NS)
    if lst is None:
        lst = el("a:effectLst", sp)
    if "reflection" in low:
        el("a:reflection", lst, blurRad=6350)
    if "soft edge" in low:
        el("a:softEdge", lst, rad=int(re.search(r"(\d+) point", low).group(1)) * 12700)
    if "glow" in low:
        el("a:glow", lst, rad=int(re.search(r"(\d+) point", low).group(1)) * 12700)
    if "shadow" in low:
        el("a:outerShdw", lst, blurRad=50800)
    if "bevel" in low:
        el("a:bevelT", el("a:sp3d", sp), prst="softRound")


@sol("pic_effect")
def _(pkg, t):
    _add_effect(_sppr(pkg, t), re.search(r"apply the (.+?) effect", t, re.I).group(1))


@sol("pic_style")
def _(pkg, t):
    name = A.norm(re.search(r"apply the (.+?) picture style", t, re.I).group(1))
    sp = _sppr(pkg, t)
    if "soft edge" in name:
        _add_effect(sp, "10 point soft edge")
    elif "shadow" in name:
        _add_effect(sp, "shadow")
    else:
        sp.find("a:prstGeom", NS).set("prst", "ellipse")


@sol("art_effect")
def _(pkg, t):
    name = re.search(r"apply the (.+?) artistic effect", t, re.I).group(1)
    blip = sec_obj(pkg, t, "pic").find(".//a:blip", NS)
    ext = el("a:extLst", blip)
    eff = etree.SubElement(el("a:ext", ext, uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}"),
                           "{http://schemas.microsoft.com/office/drawing/2010/main}imgProps")
    lay = etree.SubElement(etree.SubElement(eff, "{http://schemas.microsoft.com/office/drawing/2010/main}imgLayer"),
                           "{http://schemas.microsoft.com/office/drawing/2010/main}imgEffect")
    etree.SubElement(lay, "{http://schemas.microsoft.com/office/drawing/2010/main}" +
                     A.ARTISTIC.get(A.norm(name), "artistic" + name.title().replace(" ", "")))


@sol("pic_remove_bg")
def _(pkg, t):
    blip = sec_obj(pkg, t, "pic").find(".//a:blip", NS)
    etree.SubElement(el("a:ext", el("a:extLst", blip)),
                     "{http://schemas.microsoft.com/office/drawing/2010/main}backgroundRemoval")


@sol("pic_crop_shape")
def _(pkg, t):
    _sppr(pkg, t).find("a:prstGeom", NS).set("prst", A.SHAPES[A.norm(re.search(r"to the (.+?) shape", t).group(1))])


@sol("pic_border")
def _(pkg, t):
    m = re.search(r"Add a (.+?) pt (.+?) border", t)
    ln = el("a:ln", _sppr(pkg, t), w=int(float(m.group(1)) * 12700))
    el("a:schemeClr", el("a:solidFill", ln), val="bg1")


def _to_anchor(dr, v, h, wrap):
    if dr.tag == q("wp:inline"):
        dr.tag = q("wp:anchor")
        for k, val in dict(simplePos="0", relativeHeight="1", behindDoc="0", locked="0", layoutInCell="1",
                           allowOverlap="1").items():
            dr.set(k, val)
        ext = dr.find("wp:extent", NS)
        ext.addprevious(el("wp:simplePos", None, x=0, y=0))
        ph = el("wp:positionH", None, relativeFrom="margin")
        el("wp:align", ph).text = h
        pv = el("wp:positionV", None, relativeFrom="margin")
        el("wp:align", pv).text = v
        ext.addprevious(ph)
        ext.addprevious(pv)
        dr.find("wp:docPr", NS).addprevious(el(f"wp:wrap{wrap}", None, wrapText="bothSides"))
    else:
        for tag, val in (("wp:positionH", h), ("wp:positionV", v)):
            pos = dr.find(tag, NS)
            for x in list(pos):
                pos.remove(x)
            pos.set("relativeFrom", "margin")
            el("wp:align", pos).text = val
        for x in list(dr):
            if etree.QName(x).localname.startswith("wrap"):
                dr.replace(x, el(f"wp:wrap{wrap}", None, wrapText="bothSides"))


@sol("pic_position")
def _(pkg, t):
    v, h = A._pos_text(t)
    _to_anchor(sec_obj(pkg, t, "pic"), v, h, "Square")


@sol("pic_wrap")
def _(pkg, t):
    wrap = A.WRAP[re.search(r"to (.+?)\.$", t).group(1).lower()]
    dr = sec_obj(pkg, t, "pic")
    _to_anchor(dr, "top", "left", wrap)


@sol("callout_pos")
def _(pkg, t):
    v, h = A._pos_text(t)
    _to_anchor(A.textbox_starting(pkg.d, quotes(t)[0]), v, h, "Square")


@sol("pic_change")
def _(pkg, t):
    dr = sec_obj(pkg, t, "pic")
    rid = dr.find(".//a:blip", NS).get(q("r:embed"))
    target = pkg.d.rels("word/document.xml")[rid]
    pkg.put(target, pkg.d.files[target] + b"new")


@sol("pic_insert")
def _(pkg, t):
    target = [p for p in A.sec_paras(pkg.d, A.sec_name(t), nonempty=False) if not A.ptext(p).strip()][0]
    src = A.drawings([pkg.body], "pic")[0]
    r = el("w:r", target)
    el("w:drawing", r).append(copy.deepcopy(src))


@sol("model3d")
def _(pkg, t):
    target = [p for p in A.sec_paras(pkg.d, A.sec_name(t), nonempty=False) if not A.ptext(p).strip()][0]
    r = el("w:r", target)
    inl = el("wp:inline", el("w:drawing", r))
    el("wp:extent", inl, cx=914400, cy=914400)
    el("wp:docPr", inl, id=77, name="3D Model 1")
    el("a:graphicData", el("a:graphic", inl), uri="http://schemas.microsoft.com/office/drawing/2017/model3d")


def _textbox_xml(text, prst="rect", anchor=True, v="bottom", h="left", wrap="Tight"):
    pos = (f'<wp:positionH relativeFrom="margin"><wp:align>{h}</wp:align></wp:positionH>'
           f'<wp:positionV relativeFrom="margin"><wp:align>{v}</wp:align></wp:positionV>')
    return (f'<w:r xmlns:w="{NS["w"]}" xmlns:wp="{NS["wp"]}" xmlns:a="{NS["a"]}" xmlns:wps="{NS["wps"]}">'
            f'<w:drawing><wp:anchor behindDoc="0"><wp:simplePos x="0" y="0"/>{pos}<wp:extent cx="1" cy="1"/>'
            f'<wp:wrap{wrap} wrapText="bothSides"/><wp:docPr id="88" name="Shape 88"/><a:graphic>'
            f'<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
            f'<wps:spPr><a:prstGeom prst="{prst}"/></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>{text}'
            f'</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor>'
            f'</w:drawing></w:r>')


@sol("shape_insert")
def _(pkg, t):
    m = re.search(r"insert an? (.+?) shape that contains the text [“\"]([^”\"]+)", t)
    last = [b for b in A.blocks(pkg.d) if b.tag == A.P][-1]
    last.append(etree.fromstring(_textbox_xml(m.group(2), A.SHAPES[A.norm(m.group(1))])))


@sol("textbox_type")
def _(pkg, t):
    dr = sec_obj(pkg, t, "wps")
    run(quotes(t)[-1], dr.find(".//w:txbxContent/w:p", NS))


@sol("callout_text")
def _(pkg, t):
    dr = A.textbox_starting(pkg.d, quotes(t)[1])
    tb = dr.find(".//w:txbxContent", NS)
    run(quotes(t)[2], el("w:p", tb))


@sol("shape_style")
def _(pkg, t):
    acc = re.search(r"Accent (\d)", t).group(1)
    dr = A.textbox_starting(pkg.d, quotes(t)[-1])
    wsp = dr.find(".//wps:wsp", NS)
    style = el("wps:style", None)
    el("a:schemeClr", el("a:lnRef", style, idx=2), val=f"accent{acc}")
    wsp.find("wps:bodyPr", NS).addprevious(style)


@sol("link_textbox")
def _(pkg, t):
    boxes = A.textboxes(pkg.d)
    wsp = boxes[-1].find(".//wps:wsp", NS)
    txbx = wsp.find("wps:txbx", NS)
    txbx.addnext(el("wps:linkedTxbx", None, id=1, seq=1))
    wsp.remove(txbx)


@sol("wordart")
def _(pkg, t):
    title = A.title_para(pkg.d)
    text = A.ptext(title)
    for r in A.text_runs(title):
        title.remove(r)
    r = etree.fromstring(_textbox_xml(text))
    rp = rpr(r.find(".//w:txbxContent//w:r", NS))
    el("w14:shadow", rp)
    el("w14:textFill", rp)
    title.append(r)


# ---------- biểu đồ


def _chart_root(pkg, t):
    return pkg.x(A.chart_parts(pkg.d, A.sec_objects(pkg.d, t, "chart"))[0])


def C(tag):
    return q("c:" + tag)


@sol("chart_data")
def _(pkg, t):
    root = _chart_root(pkg, t)
    a, b = A.norm(quotes(t)[0]), A.norm(quotes(t)[1])
    new = re.search(r"to (\d+(?:\.\d+)?)", t).group(1)
    for ser in root.iter(C("ser")):
        name = A.norm(" ".join(x.text or "" for x in ser.find("c:tx", NS).iter(C("v"))))
        cats = [A.norm(x.text) for x in ser.find("c:cat", NS).iter(C("v"))]
        for s, c in ((a, b), (b, a)):
            if name == s and c in cats:
                for pt in ser.find("c:val", NS).iter(C("pt")):
                    if int(pt.get("idx")) == cats.index(c):
                        pt.find("c:v", NS).text = new


def _set_chart_type(root, name):
    tag, direction, grouping = A.CHART_TYPES[A.norm(name)]
    plot = root.find(".//c:plotArea", NS)
    old = next(x for x in plot if etree.QName(x).localname.endswith("Chart"))
    old.tag = C(tag)
    for x in old.findall("c:barDir", NS) + old.findall("c:grouping", NS):
        old.remove(x)
    if direction:
        old.insert(0, el("c:barDir", None, val=direction))
    g = el("c:grouping", None, val="standard" if tag == "lineChart" else grouping or "clustered")
    old.insert(1 if direction else 0, g)
    if grouping == "markers":
        el("c:marker", old, val=1)


@sol("chart_type")
def _(pkg, t):
    _set_chart_type(_chart_root(pkg, t), re.search(r"to an? (.+?) chart", t).group(1))


@sol("chart_title")
def _(pkg, t):
    root = _chart_root(pkg, t)
    chart = root.find("c:chart", NS)
    for x in chart.findall("c:title", NS):
        chart.remove(x)
    title = etree.Element(C("title"))
    rich = el("c:rich", el("c:tx", title))
    el("a:t", el("a:r", el("a:p", rich))).text = quotes(t)[-1]
    chart.insert(0, title)


@sol("chart_axis")
def _(pkg, t):
    ax = _chart_root(pkg, t).find(".//c:valAx", NS)
    title = etree.Element(C("title"))
    el("a:t", el("a:r", el("a:p", el("c:rich", el("c:tx", title))))).text = quotes(t)[-1]
    ax.find("c:axPos", NS).addnext(title)


@sol("chart_gridlines")
def _(pkg, t):
    for g in list(_chart_root(pkg, t).iter(C("majorGridlines"))):
        g.getparent().remove(g)


@sol("chart_legend")
def _(pkg, t):
    for g in list(_chart_root(pkg, t).iter(C("legend"))):
        g.getparent().remove(g)


@sol("chart_labels")
def _(pkg, t):
    root = _chart_root(pkg, t)
    plot = root.find(".//c:plotArea", NS)
    ch = next(x for x in plot if etree.QName(x).localname.endswith("Chart"))
    for x in ch.findall("c:dLbls", NS):
        ch.remove(x)
    dl = el("c:dLbls", None)
    el("c:dLblPos", dl, val="outEnd")
    el("c:showVal", dl, val=1)
    ch.findall("c:ser", NS)[-1].addnext(dl)


@sol("chart_switch")
def _(pkg, t):
    root = _chart_root(pkg, t)
    sers = list(root.iter(C("ser")))
    names = [" ".join(x.text for x in s.find("c:tx", NS).iter(C("v"))) for s in sers]
    cats = [x.text for x in sers[0].find("c:cat", NS).iter(C("v"))]
    vals = [[p.findtext("c:v", namespaces=NS) for p in s.find("c:val", NS).iter(C("pt"))] for s in sers]
    tmpl, parent = sers[0], sers[0].getparent()
    for s in sers:
        parent.remove(s)
    anchor = parent.find("c:varyColors", NS)
    for i, cat in enumerate(cats):
        ns = copy.deepcopy(tmpl)
        next(ns.find("c:tx", NS).iter(C("v"))).text = cat
        cache = ns.find("c:cat", NS).find(".//c:strCache", NS)
        for pt in cache.findall("c:pt", NS):
            cache.remove(pt)
        for j, n in enumerate(names):
            el("c:v", el("c:pt", cache, idx=j)).text = n
        num = ns.find("c:val", NS).find(".//c:numCache", NS)
        for pt in num.findall("c:pt", NS):
            num.remove(pt)
        for j in range(len(names)):
            el("c:v", el("c:pt", num, idx=j)).text = vals[j][i]
        anchor.addnext(ns)
        anchor = ns


@sol("chart_style", "chart_colors", "chart_layout")
def _(pkg, t):
    root = _chart_root(pkg, t)
    chart = root.find("c:chart", NS)
    if "Layout 3" in t:
        if chart.find("c:title", NS) is None:
            chart.insert(0, el("c:title", None))
        lg = chart.find("c:legend", NS)
        if lg is None:
            lg = el("c:legend", chart)
        for x in lg.findall("c:legendPos", NS):
            lg.remove(x)
        lg.insert(0, el("c:legendPos", None, val="b"))
    else:
        root.insert(0, el("c:style", None, val=7))


@sol("chart_insert")
def _(pkg, t):
    m = re.search(r"insert an? (.+?) chart", t)
    nums = re.findall(r"\(([\d.]+)\)", t)
    tag, direction, grouping = A.CHART_TYPES[A.norm(m.group(1))]
    pts = "".join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(nums))
    xml = (f'<c:chartSpace xmlns:c="{NS["c"]}" xmlns:a="{NS["a"]}"><c:chart><c:plotArea><c:{tag}>'
           f'<c:barDir val="{direction}"/><c:grouping val="{grouping}"/><c:ser><c:tx><c:strRef><c:strCache>'
           f'<c:pt idx="0"><c:v>{quotes(t)[-1]}</c:v></c:pt></c:strCache></c:strRef></c:tx><c:val><c:numRef>'
           f'<c:numCache>{pts}</c:numCache></c:numRef></c:val></c:ser></c:{tag}></c:plotArea></c:chart>'
           f'</c:chartSpace>')
    pkg.put("word/charts/chart99.xml", xml.encode())
    rid = pkg.rel("word/document.xml", RT + "chart", "charts/chart99.xml")
    target = [p for p in A.sec_paras(pkg.d, A.sec_name(t), nonempty=False) if not A.ptext(p).strip()][-1]
    inl = el("wp:inline", el("w:drawing", el("w:r", target)))
    el("wp:extent", inl, cx=5486400, cy=3200400)
    el("wp:docPr", inl, id=99, name="Chart 99")
    gd = el("a:graphicData", el("a:graphic", inl), uri="http://schemas.openxmlformats.org/drawingml/2006/chart")
    el("c:chart", gd).set(q("r:id"), rid)


# ---------- SmartArt


def _dgm_part(pkg, t, key):
    return A.dgm_parts(pkg.d, sec_obj(pkg, t, "dgm"))[key]


def _set_uid(pkg, t, key, short):
    part = _dgm_part(pkg, t, key)
    kind = {"lo": "layout", "qs": "quickstyle", "cs": "colors"}[key]
    pkg.x(part).set("uniqueId", f"urn:microsoft.com/office/officeart/2005/8/{kind}/{short}")


@sol("sa_layout")
def _(pkg, t):
    _set_uid(pkg, t, "lo", A.SA_LAYOUTS[A.norm(re.search(r"section to (.+?)\.$", t).group(1))])


@sol("sa_style")
def _(pkg, t):
    _set_uid(pkg, t, "qs", A.SA_STYLES[A.norm(re.search(r"Apply the (.+?) SmartArt style", t, re.I).group(1))])


@sol("sa_colors")
def _(pkg, t):
    _set_uid(pkg, t, "cs", A._sa_color_id(re.search(r"section to (.+?)\.$", t).group(1)))


def _dm(pkg, t):
    return pkg.x(_dgm_part(pkg, t, "dm"))


def _pt_by_text(root, text):
    for pt in root.iter(q("dgm:pt")):
        if A.norm("".join(x.text or "" for x in pt.iter(q("a:t")))) == A.norm(text):
            return pt
    raise AssertionError(text)


@sol("sa_text")
def _(pkg, t):
    pt = _pt_by_text(_dm(pkg, t), quotes(t)[1])
    ts = list(pt.iter(q("a:t")))
    ts[0].text = quotes(t)[2]
    for x in ts[1:]:
        x.text = ""


@sol("sa_add")
def _(pkg, t):
    root = _dm(pkg, t)
    after = _pt_by_text(root, quotes(t)[1])
    new = copy.deepcopy(after)
    new.set("modelId", "{NEW-1}")
    ts = list(new.iter(q("a:t")))
    ts[0].text = quotes(t)[2]
    after.addnext(new)
    cxn_list = root.find("dgm:cxnLst", NS)
    old = next(c for c in cxn_list if c.get("destId") == after.get("modelId") and c.get("type") in (None, "parOf"))
    order = int(old.get("srcOrd", 0))
    for c in cxn_list:
        if c.get("srcId") == old.get("srcId") and c.get("type") in (None, "parOf") and int(c.get("srcOrd", 0)) > order:
            c.set("srcOrd", str(int(c.get("srcOrd")) + 1))
    nc = copy.deepcopy(old)
    nc.set("modelId", "{NEW-C}")
    nc.set("destId", "{NEW-1}")
    nc.set("srcOrd", str(order + 1))
    cxn_list.append(nc)


@sol("sa_demote")
def _(pkg, t):
    root = _dm(pkg, t)
    child, parent = _pt_by_text(root, quotes(t)[1]), _pt_by_text(root, quotes(t)[2])
    for c in root.find("dgm:cxnLst", NS):
        if c.get("destId") == child.get("modelId") and c.get("type") in (None, "parOf"):
            c.set("srcId", parent.get("modelId"))
            c.set("srcOrd", "0")


@sol("sa_rtl")
def _(pkg, t):
    root = _dm(pkg, t)
    doc = next(pt for pt in root.iter(q("dgm:pt")) if pt.get("type") == "doc")
    pr = doc.find("dgm:prSet", NS)
    if pr is None:
        pr = el("dgm:prSet", doc)
    lv = pr.find("dgm:presLayoutVars", NS)
    if lv is None:
        lv = el("dgm:presLayoutVars", pr)
    for x in lv.findall("dgm:dir", NS):
        lv.remove(x)
    el("dgm:dir", lv, val="rev")


@sol("sa_effect")
def _(pkg, t):
    desc = re.search(r"apply the (.+?) shape effect", t, re.I).group(1)
    root = _dm(pkg, t)
    for pt in root.iter(q("dgm:pt")):
        if pt.get("type") in (None, "node"):
            sp = pt.find("dgm:spPr", NS)
            if sp is None:
                sp = el("dgm:spPr", pt)
            _add_effect(sp, desc if "glow" in desc.casefold() else "soft round bevel")


@sol("smartart")
def _(pkg, t):
    m = re.search(r"insert an? (.+?) SmartArt graphic", t)
    st = re.search(r"Apply the (.+?) SmartArt style", t, re.I)
    items = quotes(t)[1:]
    pts = '<dgm:pt modelId="{D}" type="doc"/>' + "".join(
        f'<dgm:pt modelId="{{P{i}}}"><dgm:t><a:p><a:r><a:t>{x}</a:t></a:r></a:p></dgm:t></dgm:pt>'
        for i, x in enumerate(items))
    cx = "".join(f'<dgm:cxn modelId="{{C{i}}}" srcId="{{D}}" destId="{{P{i}}}" srcOrd="{i}"/>'
                 for i in range(len(items)))
    ns = f'xmlns:dgm="{NS["dgm"]}" xmlns:a="{NS["a"]}"'
    parts = {
        "dm": f'<dgm:dataModel {ns}><dgm:ptLst>{pts}</dgm:ptLst><dgm:cxnLst>{cx}</dgm:cxnLst></dgm:dataModel>',
        "lo": f'<dgm:layoutDef {ns} uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/'
              f'{A.SA_LAYOUTS[A.norm(m.group(1))]}"/>',
        "qs": f'<dgm:styleDef {ns} uniqueId="urn:microsoft.com/office/officeart/2005/8/quickstyle/'
              f'{A.SA_STYLES[A.norm(st.group(1))]}"/>',
        "cs": f'<dgm:colorsDef {ns} uniqueId="urn:microsoft.com/office/officeart/2005/8/colors/accent1_2"/>',
    }
    ids = {}
    for k, xml in parts.items():
        name = f"diagrams/{k}99.xml"
        pkg.put("word/" + name, xml.encode())
        ids[k] = pkg.rel("word/document.xml", RT + "diagram" + k, name)
    target = [p for p in A.sec_paras(pkg.d, A.sec_name(t), nonempty=False) if not A.ptext(p).strip()][-1]
    inl = el("wp:inline", el("w:drawing", el("w:r", target)))
    el("wp:extent", inl, cx=5486400, cy=3200400)
    el("wp:docPr", inl, id=98, name="Diagram 98")
    gd = el("a:graphicData", el("a:graphic", inl), uri="http://schemas.openxmlformats.org/drawingml/2006/diagram")
    rel = el("dgm:relIds", gd)
    for k, rid in ids.items():
        rel.set(q("r:" + k), rid)


# ---------- cộng tác


def _comments_root(pkg):
    return pkg.x("word/comments.xml", f'<w:comments xmlns:w="{NS["w"]}"/>')


@sol("comment_add")
def _(pkg, t):
    phrase, body = quotes(t)[1], quotes(t)[2]
    p = [x for x in A.sec_paras(pkg.d, A.sec_name(t)) if phrase in A.ptext(x)][0]
    runs = split_phrase(p, phrase)
    s, e = etree.Element(W + "commentRangeStart"), etree.Element(W + "commentRangeEnd")
    wattr(s, "id", 50)
    wattr(e, "id", 50)
    runs[0].addprevious(s)
    runs[-1].addnext(e)
    wattr(el("w:commentReference", el("w:r", None)), "id", 50)
    e.addnext(etree.Element(W + "r"))
    wattr(el("w:commentReference", e.getnext()), "id", 50)
    c = el("w:comment", _comments_root(pkg))
    wattr(c, "id", 50)
    run(body, el("w:p", c))


def _drop_comment(pkg, cid):
    for x in list(pkg.body.iter(W + "commentRangeStart", W + "commentRangeEnd", W + "commentReference")):
        if x.get(W + "id") == cid:
            parent = x.getparent()
            parent.remove(x)
            if parent.tag == W + "r" and len(parent) <= 1:
                parent.getparent().remove(parent)
    root = _comments_root(pkg)
    for c in root.findall("w:comment", NS):
        if c.get(W + "id") == cid:
            root.remove(c)


@sol("comment_delete")
def _(pkg, t):
    phrase = quotes(t)[-1]
    for cid, txt in A.comment_ranges(pkg.d).items():
        if A.norm(phrase) in A.norm(txt):
            _drop_comment(pkg, cid)


@sol("comment_delete_all")
def _(pkg, t):
    for cid in list(A.comments(pkg.d)):
        _drop_comment(pkg, cid)


@sol("comment_reply")
def _(pkg, t):
    reply = quotes(t)[-1]
    root = _comments_root(pkg)
    c = el("w:comment", root)
    wattr(c, "id", 60)
    p = el("w:p", c)
    p.set(q("w14:paraId"), "0000AAAA")
    run(reply, p)
    ext = pkg.x("word/commentsExtended.xml", f'<w15:commentsEx xmlns:w15="{NS["w15"]}"/>')
    el("w15:commentEx", ext, w15_paraId="0000AAAA", w15_paraIdParent="0000BBBB", w15_done=0)


@sol("comment_resolve")
def _(pkg, t):
    ext = pkg.x("word/commentsExtended.xml", f'<w15:commentsEx xmlns:w15="{NS["w15"]}"/>')
    el("w15:commentEx", ext, w15_paraId="0000CCCC", w15_done=1)


@sol("track_on")
def _(pkg, t):
    old, new = quotes(t)[0], quotes(t)[1]
    el("w:trackRevisions", pkg.settings())
    p = next(x for x in A.all_paras(pkg.body) if re.search(rf"\b{re.escape(old)}\b", A.ptext(x)))
    runs = split_phrase(p, old)
    d = etree.Element(W + "del")
    runs[0].addprevious(d)
    for r in runs:
        d.append(r)
        r.find(A.T).tag = W + "delText"
    ins = etree.Element(W + "ins")
    run(new, ins)
    d.addnext(ins)


def _accept_all(pkg, keep_ins=True, keep_del=False, only=None):
    for ins in list(pkg.body.iter(W + "ins")):
        txt = "".join(x.text or "" for x in ins.iter(A.T))
        keep = keep_ins if only is None else A.norm(txt) == A.norm(only)
        parent = ins.getparent()
        if keep:
            for ch in list(ins):
                ins.addprevious(ch)
        parent.remove(ins)
    for d in list(pkg.body.iter(W + "del")):
        parent = d.getparent()
        if keep_del or only is not None:
            for ch in list(d):
                for dt in ch.iter(W + "delText"):
                    dt.tag = A.T
                d.addprevious(ch)
        parent.remove(d)
    for ch in list(pkg.body.iter(W + "rPrChange", W + "pPrChange")):
        old = ch.find(W + "rPr")
        pr = ch.getparent()
        if pr.tag == W + "rPr" and old is not None:
            r = pr.getparent()
            r.replace(pr, old)
        else:
            pr.remove(ch)


@sol("track_accept")
def _(pkg, t):
    _accept_all(pkg)


@sol("track_reject")
def _(pkg, t):
    _accept_all(pkg, only=quotes(t)[0])


@sol("mailmerge")
def _(pkg, t):
    src = re.search(r"Use the (\S+\.xlsx)", t).group(1)
    st = pkg.settings()
    mm = el("w:mailMerge", st)
    wattr(el("w:mainDocumentType", mm), "val", "formLetters")
    wattr(el("w:query", mm), "val", f"SELECT * FROM `{src}`")
    p = para_with(pkg, "Dear", start=True)
    _field(p, "MERGEFIELD First_Name", "«First_Name»")


@sol("content_control")
def _(pkg, t):
    p = [x for x in A.all_paras(pkg.body) if A.ptext(x).strip() == "Name:"][-1]
    sdt = el("w:sdt", p)
    pr = el("w:sdtPr", sdt)
    wattr(el("w:alias", pr), "val", quotes(t)[-1])
    el("w:text", pr)
    run("Click here", el("w:sdtContent", sdt))


@sol("cover_page")
def _(pkg, t):
    sdt = etree.Element(W + "sdt")
    gal = el("w:docPartGallery", el("w:docPartObj", el("w:sdtPr", sdt)))
    wattr(gal, "val", "Cover Pages")
    run("Cover", el("w:p", el("w:sdtContent", sdt)))
    pkg.body.insert(0, sdt)


@sol("macro")
def _(pkg, t):
    with zipfile.ZipFile(pkg.path.with_suffix(".docm"), "w") as z:
        z.writestr("word/vbaProject.bin", b"vba")


@sol("encrypt")
def _(pkg, t):
    pkg.encrypted = True


def solve(path: Path, dang: str, text: str) -> None:
    pkg = Pkg(path)
    SOL[dang](pkg, text)
    if getattr(pkg, "encrypted", False):
        Path(path).write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 504)
        return
    pkg.save()
