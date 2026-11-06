#!/usr/bin/env python3
"""Excel の図形（xl/drawings/drawingN.xml）を SVG / PNG に描いて、見た目を確かめる。

使い方:
  python3 render_drawing.py 図.xlsx 出力.png [--drawing xl/drawings/drawing1.xml]
                            [--y0 上端pt] [--y1 下端pt] [--scale 1.0]

- Excel も LibreOffice も無い Mac で、図の配置・重なり・はみ出しを見るための近似の描画。
- まとまり（グループ）・回転・反転・カギ線・吹き出し・ハイパーリンクの札も描く。
- 字の幅は近似なので、最後は Excel か Numbers で開いて確かめる。
- PNG にするには rsvg-convert（brew の librsvg）を使う。出力が .svg なら SVG のまま書く。
"""
import argparse
import html
import math
import re
import subprocess
import sys
import zipfile

EMU = 12700  # 1pt
THEME = {
    "dk1": "000000", "lt1": "FFFFFF", "dk2": "0E2841", "lt2": "E8E8E8",
    "tx1": "000000", "bg1": "FFFFFF", "tx2": "0E2841", "bg2": "E8E8E8",
    "accent1": "156082", "accent2": "E97132", "accent3": "196B24",
    "accent4": "0F9ED5", "accent5": "A02B93", "accent6": "4EA72E",
}


def load_theme(z):
    """ブックの theme1.xml から色を読む（無ければ既定の Office の色）。"""
    try:
        t = z.read("xl/theme/theme1.xml").decode("utf-8")
    except KeyError:
        return
    for name in ["dk1", "lt1", "dk2", "lt2"] + [f"accent{i}" for i in range(1, 7)]:
        m = re.search(rf"<a:{name}>.*?(?:srgbClr val|lastClr)=\"([0-9A-Fa-f]{{6}})\"", t, re.S)
        if m:
            THEME[name] = m.group(1).upper()
    THEME.update(tx1=THEME["dk1"], bg1=THEME["lt1"], tx2=THEME["dk2"], bg2=THEME["lt2"])


def apply_mods(rgb, mods):
    r, g, b = rgb
    lm = re.search(r'lumMod val="(\d+)"', mods)
    lo = re.search(r'lumOff val="(\d+)"', mods)
    sh = re.search(r'shade val="(\d+)"', mods)
    ti = re.search(r'tint val="(\d+)"', mods)
    if lm or lo:
        f = int(lm.group(1)) / 1e5 if lm else 1
        o = int(lo.group(1)) / 1e5 if lo else 0
        r, g, b = [min(255, int(c * f + 255 * o)) for c in (r, g, b)]
    if sh:
        f = int(sh.group(1)) / 1e5
        r, g, b = [int(c * f) for c in (r, g, b)]
    if ti:
        f = int(ti.group(1)) / 1e5
        r, g, b = [int(255 - (255 - c) * f) for c in (r, g, b)]
    return r, g, b


def color_of(node):
    """色の指定 1 つ（solidFill の中など）を #RRGGBB にする。読めなければ None。"""
    if not node:
        return None
    m = re.search(r'<a:srgbClr val="([0-9A-Fa-f]{6})"\s*(/>|>(.*?)</a:srgbClr>)', node, re.S)
    if m:
        h = m.group(1)
        rgb = apply_mods(tuple(int(h[i:i + 2], 16) for i in (0, 2, 4)), m.group(3) or "")
        return "#%02X%02X%02X" % rgb
    m = re.search(r'<a:schemeClr val="(\w+)"\s*(/>|>(.*?)</a:schemeClr>)', node, re.S)
    if m:
        h = THEME.get(m.group(1), "888888")
        rgb = apply_mods(tuple(int(h[i:i + 2], 16) for i in (0, 2, 4)), m.group(3) or "")
        return "#%02X%02X%02X" % rgb
    m = re.search(r'<a:prstClr val="(\w+)"', node)
    if m:
        return {"black": "#000000", "white": "#FFFFFF", "red": "#FF0000"}.get(m.group(1), "#000000")
    m = re.search(r'<a:scrgbClr r="(\d+)" g="(\d+)" b="(\d+)"', node)
    if m:
        return "#%02X%02X%02X" % tuple(int(int(v) / 1e5 * 255) for v in m.groups())
    return None


def first(pattern, text, flags=re.S):
    m = re.search(pattern, text, flags)
    return m.group(1) if m else None


class Xform:
    """グループの座標変換（子の座標 → シート上の pt）。"""

    def __init__(self, ox=0, oy=0, sx=1, sy=1, cx=0, cy=0):
        self.ox, self.oy, self.sx, self.sy, self.cx, self.cy = ox, oy, sx, sy, cx, cy

    def pt(self, x, y):
        return (self.ox + (x - self.cx) * self.sx) / EMU, (self.oy + (y - self.cy) * self.sy) / EMU

    def scale(self, w, h):
        return w * self.sx / EMU, h * self.sy / EMU

    def child(self, grp_xfrm):
        off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"', grp_xfrm)
        ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', grp_xfrm)
        choff = re.search(r'<a:chOff x="(-?\d+)" y="(-?\d+)"', grp_xfrm)
        chext = re.search(r'<a:chExt cx="(\d+)" cy="(\d+)"', grp_xfrm)
        ox, oy = (int(v) for v in off.groups())
        ew, eh = (int(v) for v in ext.groups())
        cx, cy = (int(v) for v in choff.groups())
        cw, ch = (int(v) for v in chext.groups())
        sx = ew / cw if cw else 1
        sy = eh / ch if ch else 1
        pox = self.ox + (ox - self.cx) * self.sx
        poy = self.oy + (oy - self.cy) * self.sy
        return Xform(pox, poy, self.sx * sx, self.sy * sy, cx, cy)


def find_close(xml, start, tag):
    depth = 0
    pat = re.compile(rf"<{tag}[\s>]|</{tag}>")
    for m in pat.finditer(xml, start):
        if m.group(0).startswith("</"):
            depth -= 1
            if depth == 0:
                return m.end()
        else:
            depth += 1
    return len(xml)


def split_children(xml):
    """xml の中の図形（sp / cxnSp / grpSp / pic）を、入れ子の外側から順に返す。"""
    out = []
    pat = re.compile(r"<(xdr:sp|xdr:cxnSp|xdr:grpSp|xdr:pic)[\s>]")
    i = 0
    while True:
        m = pat.search(xml, i)
        if not m:
            return out
        tag = m.group(1)
        end = find_close(xml, m.start(), tag)
        out.append((tag, xml[m.start():end]))
        i = end


def text_width(s, sz):
    w = 0
    for ch in s:
        o = ord(ch)
        if o < 0x7F:
            w += sz * (0.33 if ch in " il.,:;|!'()[]" else 0.56)
        elif 0xFF61 <= o <= 0xFF9F:
            w += sz * 0.5
        else:
            w += sz
    return w


class Svg:
    def __init__(self):
        self.items = []
        self.shapes = []  # --list 用: (種類, x, y, w, h, 文字)
        self.bounds = [1e9, 1e9, -1e9, -1e9]

    def grow(self, x0, y0, x1, y1):
        b = self.bounds
        b[0], b[1] = min(b[0], x0, x1), min(b[1], y0, y1)
        b[2], b[3] = max(b[2], x0, x1), max(b[3], y0, y1)

    def add(self, s):
        self.items.append(s)


def arrow(svg, x1, y1, x2, y2, c, w):
    a = math.atan2(y2 - y1, x2 - x1)
    L = 6 + w * 2
    pts = [(x2, y2), (x2 - L * math.cos(a - 0.45), y2 - L * math.sin(a - 0.45)),
           (x2 - L * math.cos(a + 0.45), y2 - L * math.sin(a + 0.45))]
    svg.add('<polygon points="%s" fill="%s"/>' % (" ".join("%.1f,%.1f" % p for p in pts), c))


def line_style(sppr, style):
    ln = first(r"(<a:ln\b[^>]*/>|<a:ln\b.*?</a:ln>)", sppr) or ""
    if "<a:noFill/>" in ln:
        return None
    w = first(r'<a:ln w="(\d+)"', ln)
    w = int(w) / EMU if w else 0.75
    c = color_of(first(r"<a:solidFill>(.*?)</a:solidFill>", ln) or "")
    if c is None:
        idx = first(r'<a:lnRef idx="(\d+)"', style or "")
        if idx and idx != "0":
            c = color_of(first(r"(<a:lnRef.*?</a:lnRef>)", style)) or "#000000"
            if not first(r'<a:ln w="(\d+)"', ln):
                w = {"1": 0.75, "2": 1.0, "3": 1.5}.get(idx, 1.0)
        elif ln:
            c = "#000000"
        else:
            return None
    dash = ""
    d = first(r'<a:prstDash val="(\w+)"', ln)
    if d and d != "solid":
        dash = ' stroke-dasharray="%s"' % ("2 2" if "sysDot" in d or d == "dot" else "6 3")
    return c, w, dash, ln


def fill_of(sppr, style):
    if re.search(r"</a:(?:prstGeom|custGeom)>\s*<a:noFill/>", sppr):
        return "none"
    f = first(r"</a:(?:prstGeom|custGeom)>\s*<a:solidFill>(.*?)</a:solidFill>", sppr)
    if f:
        return color_of(f) or "none"
    if re.search(r"</a:(?:prstGeom|custGeom)>\s*<a:gradFill", sppr):
        return color_of(first(r"<a:gs\b[^>]*>(.*?)</a:gs>", sppr) or "") or "none"
    idx = first(r'<a:fillRef idx="(\d+)"', style or "")
    if idx and idx != "0":
        return color_of(first(r"(<a:fillRef.*?</a:fillRef>)", style)) or "none"
    return "none"


def rotate_points(pts, cx, cy, deg):
    a = math.radians(deg)
    return [(cx + (px - cx) * math.cos(a) - (py - cy) * math.sin(a),
             cy + (px - cx) * math.sin(a) + (py - cy) * math.cos(a)) for px, py in pts]


def draw_connector(svg, sppr, prst, x, y, w, h, rot, flipH, flipV, ls):
    stroke, sw, dash, ln = ls
    x1, x2 = (x + w, x) if flipH else (x, x + w)
    y1, y2 = (y + h, y) if flipV else (y, y + h)
    if prst.startswith("bentConnector"):
        adj = int(first(r'name="adj1" fmla="val (-?\d+)"', sppr) or 50000) / 1e5
        if prst == "bentConnector2":
            pts = [(x1, y1), (x2, y1), (x2, y2)]
        else:
            mx = x1 + (x2 - x1) * adj
            pts = [(x1, y1), (mx, y1), (mx, y2), (x2, y2)]
    else:
        pts = [(x1, y1), (x2, y2)]
    if rot:
        pts = rotate_points(pts, x + w / 2, y + h / 2, rot)
    svg.add('<polyline points="%s" fill="none" stroke="%s" stroke-width="%.2f"%s/>' % (
        " ".join("%.1f,%.1f" % p for p in pts), stroke, sw, dash))
    if re.search(r'<a:tailEnd type="(?!none)', ln):
        (ax1, ay1), (ax2, ay2) = pts[-2], pts[-1]
        arrow(svg, ax1, ay1, ax2, ay2, stroke, sw)
    if re.search(r'<a:headEnd type="(?!none)', ln):
        (bx1, by1), (bx2, by2) = pts[1], pts[0]
        arrow(svg, bx1, by1, bx2, by2, stroke, sw)


def draw_shape(svg, xml, xf, links):
    sppr = first(r"(<xdr:spPr\b.*?</xdr:spPr>)", xml) or ""
    style = first(r"(<xdr:style>.*?</xdr:style>)", xml) or ""
    off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/>\s*<a:ext cx="(\d+)" cy="(\d+)"', sppr)
    if not off:
        return
    x, y = xf.pt(int(off.group(1)), int(off.group(2)))
    w, h = xf.scale(int(off.group(3)), int(off.group(4)))
    rot = int(first(r'<a:xfrm[^>]*rot="(-?\d+)"', sppr) or 0) / 60000
    flipH = bool(re.search(r'<a:xfrm[^>]*flipH="1"', sppr))
    flipV = bool(re.search(r'<a:xfrm[^>]*flipV="1"', sppr))
    prst = first(r'prst="(\w+)"', sppr) or "rect"
    ls = line_style(sppr, style)
    stroke, sw, dash = (ls[0], ls[1], ls[2]) if ls else ("none", 0, "")
    svg.grow(x, y, x + w, y + h)
    label = html.unescape("".join(re.findall(r"<a:t>(.*?)</a:t>", xml, re.S)))
    svg.shapes.append((prst, x, y, w, h, label))
    if xml.startswith("<xdr:cxnSp") or prst in ("line", "straightConnector1") or \
            prst.startswith(("bentConnector", "curvedConnector")):
        if ls:
            draw_connector(svg, sppr, prst, x, y, w, h, rot, flipH, flipV, ls)
        return
    tr = ' transform="rotate(%.1f %.1f %.1f)"' % (rot, x + w / 2, y + h / 2) if rot else ""
    fill = fill_of(sppr, style)
    common = 'fill="%s" stroke="%s" stroke-width="%.2f"%s%s' % (fill, stroke, sw, dash, tr)
    if prst == "roundRect":
        adj = int(first(r'name="adj" fmla="val (\d+)"', sppr) or 16667) / 1e5
        svg.add('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" rx="%.1f" %s/>' % (
            x, y, w, h, min(w, h) * adj, common))
    elif prst == "ellipse":
        svg.add('<ellipse cx="%.1f" cy="%.1f" rx="%.1f" ry="%.1f" %s/>' % (x + w / 2, y + h / 2, w / 2, h / 2, common))
    elif prst in ("homePlate", "chevron"):
        t = min(w, h) * int(first(r'name="adj" fmla="val (\d+)"', sppr) or 50000) / 1e5
        p = [(x, y), (x + w - t, y), (x + w, y + h / 2), (x + w - t, y + h), (x, y + h)]
        if prst == "chevron":
            p.append((x + t, y + h / 2))
        svg.add('<polygon points="%s" %s/>' % (" ".join("%.1f,%.1f" % q for q in p), common))
    elif prst == "plus":
        t = min(w, h) * int(first(r'name="adj" fmla="val (\d+)"', sppr) or 25000) / 1e5
        d = "M%.1f,%.1f h%.1f v%.1f h%.1f v%.1f h-%.1f v%.1f h-%.1f v-%.1f h-%.1f v-%.1f h%.1f z" % (
            x + t, y, w - 2 * t, t, t, h - 2 * t, t, t, w - 2 * t, t, t, h - 2 * t, t)
        svg.add('<path d="%s" %s/>' % (d, common))
    elif prst.startswith(("borderCallout", "callout", "accentCallout")):
        svg.add('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" %s/>' % (x, y, w, h, common))
        a = [int(first(rf'name="adj{i}" fmla="val (-?\d+)"', sppr) or d) / 1e5
             for i, d in zip(range(1, 5), (18750, -8333, 112500, -38333))]
        sy1, sx1, ey1, ex1 = y + h * a[0], x + w * a[1], y + h * a[2], x + w * a[3]
        svg.add('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="%s" stroke-width="%.2f"/>' % (
            sx1, sy1, ex1, ey1, stroke if stroke != "none" else "#000", max(sw, .75)))
        svg.grow(ex1, ey1, ex1, ey1)
    else:
        svg.add('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" %s/>' % (x, y, w, h, common))
    draw_text(svg, xml, x, y, w, h, style, links, tr)


def layout_lines(tb, w, ins, wrap, default_col):
    lines = []
    for m in re.finditer(r"<a:p>(.*?)</a:p>|<a:p/>", tb, re.S):
        pa = m.group(1) or ""
        algn = first(r'<a:pPr[^>]*algn="(\w+)"', pa) or "l"
        segs = []
        for r in re.findall(r"<a:r>.*?</a:r>|<a:br\b[^>]*/>", pa, re.S):
            if r.startswith("<a:br"):
                segs.append(None)
                continue
            rpr = first(r"(<a:rPr\b[^>]*?(?:/>|>.*?</a:rPr>))", r) or ""
            sz = int(first(r'sz="(\d+)"', rpr) or first(r'<a:endParaRPr[^>]*sz="(\d+)"', pa) or 1100) / 100
            b = first(r'\bb="(\d)"', rpr) == "1"
            col = color_of(first(r"<a:solidFill>(.*?)</a:solidFill>", rpr) or "") or default_col
            u = bool(re.search(r'u="sng"', rpr))
            t = html.unescape("".join(re.findall(r"<a:t>(.*?)</a:t>", r, re.S)))
            segs.append((t, sz, b, col, u))
        if not segs:
            sz = int(first(r'sz="(\d+)"', pa) or 1100) / 100
            lines.append((algn, sz, []))
            continue
        maxw = w - ins[0] - ins[2] if wrap != "none" else 1e9
        cur, curw, cursz = [], 0.0, 0
        for s in segs:
            if s is None:
                lines.append((algn, cursz or 11, cur))
                cur, curw, cursz = [], 0.0, 0
                continue
            t, sz, b, col, u = s
            buf = ""
            for ch in t:
                cw = text_width(ch, sz)
                if curw + cw > maxw and (cur or buf):
                    if buf:
                        cur.append((buf, sz, b, col, u))
                    lines.append((algn, max(cursz, sz), cur))
                    cur, curw, cursz, buf = [], 0.0, 0, ""
                buf += ch
                curw += cw
                cursz = max(cursz, sz)
            if buf:
                cur.append((buf, sz, b, col, u))
        lines.append((algn, cursz or 11, cur))
    return lines


def draw_text(svg, xml, x, y, w, h, style, links, tr):
    tb = first(r"<xdr:txBody>(.*?)</xdr:txBody>", xml)
    if not tb:
        return
    body = first(r"(<a:bodyPr\b[^>]*?(?:/>|>.*?</a:bodyPr>))", tb) or ""
    anchor = first(r'anchor="(\w+)"', body) or "t"
    wrap = first(r'wrap="(\w+)"', body) or "square"
    ins = [first(rf'{k}Ins="(\d+)"', body) for k in "ltrb"]
    ins = [int(v) / EMU if v else d for v, d in zip(ins, (7.2, 3.6, 7.2, 3.6))]
    default_col = color_of(first(r"(<a:fontRef.*?</a:fontRef>)", style)) or "#000000"
    is_link = bool(re.search(r'<a:hlinkClick r:id="\w+"', xml))
    lines = layout_lines(tb, w, ins, wrap, default_col)
    total = sum(l[1] * 1.3 for l in lines)
    if anchor == "ctr":
        ty = y + (h - total) / 2
    elif anchor == "b":
        ty = y + h - ins[3] - total
    else:
        ty = y + ins[1]
    out = []
    for algn, sz, segs in lines:
        ty += sz * 1.3
        lw = sum(text_width(s[0], s[1]) for s in segs)
        if algn == "ctr":
            tx = x + ins[0] + (w - ins[0] - ins[2] - lw) / 2
        elif algn == "r":
            tx = x + w - ins[2] - lw
        else:
            tx = x + ins[0]
        for t, s, b, col, u in segs:
            if is_link:
                u = True
            deco = ' text-decoration="underline"' if u else ""
            out.append('<text x="%.1f" y="%.1f" font-size="%.1f" fill="%s"%s%s>%s</text>' % (
                tx, ty - s * 0.28, s, col, ' font-weight="bold"' if b else "", deco, html.escape(t)))
            svg.grow(tx, ty - s, tx + text_width(t, s), ty)
            tx += text_width(t, s)
    svg.add("<g%s>%s</g>" % (tr, "".join(out)) if tr else "".join(out))


def walk(svg, xml, xf, links):
    for tag, part in split_children(xml):
        if tag == "xdr:grpSp":
            gx = first(r"<xdr:grpSpPr>\s*(<a:xfrm.*?</a:xfrm>)", part)
            start = part.find("</xdr:grpSpPr>")
            inner = part[start + len("</xdr:grpSpPr>"):part.rfind("</xdr:grpSp>")]
            walk(svg, inner, xf.child(gx) if gx else xf, links)
        elif tag in ("xdr:sp", "xdr:cxnSp"):
            draw_shape(svg, part, xf, links)
        elif tag == "xdr:pic":
            off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/>\s*<a:ext cx="(\d+)" cy="(\d+)"', part)
            if off:
                x, y = xf.pt(int(off.group(1)), int(off.group(2)))
                w, h = xf.scale(int(off.group(3)), int(off.group(4)))
                svg.add('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="#EEE" stroke="#999"/>'
                        '<text x="%.1f" y="%.1f" font-size="10" fill="#666">（画像）</text>' % (
                            x, y, w, h, x + 4, y + 12))
                svg.grow(x, y, x + w, y + h)


def render(xlsx, drawing="xl/drawings/drawing1.xml", y0=None, y1=None):
    z = zipfile.ZipFile(xlsx)
    load_theme(z)
    xml = z.read(drawing).decode("utf-8")
    svg = Svg()
    anchors = re.findall(r"<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b.*?</xdr:\1>", xml, re.S)
    for m in re.finditer(r"<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b.*?</xdr:\1>", xml, re.S):
        walk(svg, m.group(0), Xform(), {})
    bx0, by0, bx1, by1 = svg.bounds
    bx0 = min(0, bx0) - 10
    by0 = y0 if y0 is not None else by0 - 10
    by1 = y1 if y1 is not None else by1 + 10
    bx1 += 10
    W, H = bx1 - bx0, by1 - by0
    head = ('<svg xmlns="http://www.w3.org/2000/svg" width="%.0f" height="%.0f" viewBox="%.1f %.1f %.1f %.1f" '
            'font-family="Hiragino Sans, sans-serif">'
            '<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="white"/>' % (W, H, bx0, by0, W, H, bx0, by0, W, H))
    render.shapes = svg.shapes
    return head + "".join(svg.items) + "</svg>", (W, H), len(anchors)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("xlsx")
    ap.add_argument("out", help=".svg か .png")
    ap.add_argument("--drawing", default="xl/drawings/drawing1.xml")
    ap.add_argument("--y0", type=float)
    ap.add_argument("--y1", type=float)
    ap.add_argument("--scale", type=float, default=1.0, help="PNG の倍率")
    ap.add_argument("--list", action="store_true", help="図形の種類・位置(pt)・文字を一覧で出す")
    a = ap.parse_args()
    svg, (W, H), n = render(a.xlsx, a.drawing, a.y0, a.y1)
    if a.list:
        for prst, x, y, w, h, label in render.shapes:
            if (a.y0 is None or y + h >= a.y0) and (a.y1 is None or y <= a.y1):
                print("%-18s x=%7.1f y=%7.1f w=%7.1f h=%6.1f  %s" % (prst, x, y, w, h, label[:60]))
    if a.out.endswith(".svg"):
        open(a.out, "w", encoding="utf-8").write(svg)
    else:
        tmp = a.out + ".svg"
        open(tmp, "w", encoding="utf-8").write(svg)
        subprocess.run(["rsvg-convert", "-z", str(a.scale), "-o", a.out, tmp], check=True)
    print("%s (%.0f x %.0f pt, anchors=%d)" % (a.out, W, H, n))


if __name__ == "__main__":
    sys.exit(main())
