"""Excel の図形（DrawingML）を台本で組み立て、既存の xlsx の図に足す道具。

- 座標と大きさは pt（Excel の 100% 表示での 1 点）。
- 矢印は箱のつなぎ目（connection site）につなげる。つないだ矢印は、人が箱を動かすと付いてくる。
  つなぎ目の番号: 四角・角丸・ホームベース は 0=上 1=左 2=下 3=右。
  ホームベース（homePlate）の上下のつなぎ目は、矢じり部分を除いた四角の真ん中にある。
- 足す図形は全体で 1 つのまとまり（グループ）にし、その中も意味ごとに小さなまとまりにできる。
- まとまりには目印（tag）を付ける。同じ目印で書き直すと、前に足した物とそのリンクを消してから足す。
- ハイパーリンクは図形そのものに付ける。リンク先は drawing の rels に足す。
- 既存の図形・セル・ほかの部品は 1 バイトも変えない。

色は "RRGGBB" か、テーマの色 ("accent6", 明るさ%, 足す白%)。例: ("accent6", 20, 80) は薄い緑。
"""
import html
import re
import zipfile
from contextlib import contextmanager

EMU = 12700
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
HLINK_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"


def esc(s):
    return html.escape(str(s), quote=False)


def clr(c):
    if isinstance(c, tuple):
        name, lm, lo = (list(c) + [None, None])[:3]
        mods = (f'<a:lumMod val="{lm * 1000}"/>' if lm is not None else "") + \
               (f'<a:lumOff val="{lo * 1000}"/>' if lo is not None else "")
        return f'<a:schemeClr val="{name}">{mods}</a:schemeClr>' if mods else f'<a:schemeClr val="{name}"/>'
    return f'<a:srgbClr val="{c}"/>'


def emu(v):
    return int(round(v * EMU))


class Shape:
    """足した図形の控え。矢印をつなぐときに使う。"""

    def __init__(self, local_id, prst, x, y, w, h, adj):
        self.local_id, self.prst = local_id, prst
        self.x, self.y, self.w, self.h, self.adj = x, y, w, h, adj

    @property
    def cx(self):
        return self.x + self.w / 2

    @property
    def cy(self):
        return self.y + self.h / 2

    def site(self, idx):
        """つなぎ目 idx の位置（pt）。"""
        x, y, w, h = self.x, self.y, self.w, self.h
        if self.prst == "homePlate":
            dx1 = min(w, h) * (self.adj if self.adj is not None else 50000) / 1e5
            x2 = x + (w - dx1) / 2
            return [(x2, y), (x, y + h / 2), (x2, y + h), (x + w, y + h / 2)][idx]
        if self.prst in ("rect", "roundRect", "flowChartProcess", "snipRoundRect", "foldedCorner", "ellipse", "flowChartPredefinedProcess"):
            return [(x + w / 2, y), (x, y + h / 2), (x + w / 2, y + h), (x + w, y + h / 2)][idx]
        raise ValueError(f"{self.prst} のつなぎ目はこの道具では扱っていない")


class _Group:
    def __init__(self, name):
        self.name = name
        self.items = []  # ("leaf", xml, x, y, w, h) or ("group", _Group)


class Drawing:
    """図形を集めて、1 つのまとまりの XML にする。"""

    def __init__(self, tag, name="まとまり"):
        self.tag = tag            # 書き直すときの目印（cNvPr の descr に入る）
        self.root = _Group(name)
        self.stack = [self.root]
        self.links = []           # (rid, url)
        self.count = 0            # 仮の id の通し番号（書き込むときに本当の id にする）

    # ---- 内部 -------------------------------------------------------------
    def _new_id(self):
        self.count += 1
        return self.count

    @staticmethod
    def _ph(n):
        return f"@ID{n}@"

    def _rpr(self, o):
        u = ' u="sng"' if o.get("underline") else ""
        font = o.get("font")
        fx = f'<a:latin typeface="{font}"/><a:ea typeface="{font}"/>' if font else ""
        return (f'<a:rPr lang="ja-JP" altLang="en-US" sz="{int(o.get("size", 11) * 100)}" '
                f'b="{1 if o.get("bold") else 0}" i="0"{u}>'
                f'<a:solidFill>{clr(o.get("color", "222222"))}</a:solidFill>{fx}</a:rPr>')

    def _paras(self, text, base):
        """text は 文字列（\\n で段落）か、段落の list。段落は 文字列 か [(文字列, {上書き}), ...]。"""
        if isinstance(text, str):
            text = text.split("\n")
        out = []
        for p in text:
            if isinstance(p, str):
                p = [(p, {})]
            elif isinstance(p, tuple):
                p = [p]
            out.append([(t, {**base, **o}) for t, o in p])
        return out

    def _txbody(self, text, base, align, anchor, ins, wrap):
        body = [f'<xdr:txBody><a:bodyPr vertOverflow="overflow" horzOverflow="overflow" wrap="{wrap}" '
                f'lIns="{emu(ins[0])}" tIns="{emu(ins[1])}" rIns="{emu(ins[2])}" bIns="{emu(ins[3])}" '
                f'rtlCol="0" anchor="{anchor}"><a:noAutofit/></a:bodyPr><a:lstStyle/>']
        for runs in self._paras(text, base):
            body.append(f'<a:p><a:pPr algn="{align}"/>')
            for t, o in runs:
                body.append(f'<a:r>{self._rpr(o)}<a:t>{esc(t)}</a:t></a:r>')
            o = runs[-1][1] if runs else base
            body.append(f'<a:endParaRPr lang="ja-JP" altLang="en-US" sz="{int(o.get("size", 11) * 100)}"/></a:p>')
        body.append("</xdr:txBody>")
        return "".join(body)

    def _add(self, xml, x, y, w, h):
        self.stack[-1].items.append(("leaf", xml, x, y, w, h))

    # ---- まとまり ---------------------------------------------------------
    @contextmanager
    def group(self, name):
        """with d.group("問題点1"): の中で足した図形を、小さなまとまりにする。入れ子にできる。"""
        g = _Group(name)
        self.stack[-1].items.append(("group", g))
        self.stack.append(g)
        try:
            yield g
        finally:
            self.stack.pop()

    # ---- 図形 -------------------------------------------------------------
    def shape(self, x, y, w, h, prst="rect", fill=None, line=None, lw=0.75, dash=None,
              adj=None, adjs=None, text=None, size=11, color="222222", bold=False, font=None,
              align="ctr", anchor="ctr", ins=(4, 2, 4, 2), wrap="square", rot=0, link=None,
              name="図形"):
        n = self._new_id()
        gds = f'<a:gd name="adj" fmla="val {adj}"/>' if adj is not None else ""
        if adjs:
            gds = "".join(f'<a:gd name="{k}" fmla="val {int(v)}"/>' for k, v in adjs.items())
        fillx = f"<a:solidFill>{clr(fill)}</a:solidFill>" if fill else "<a:noFill/>"
        if line:
            linex = (f'<a:ln w="{emu(lw)}"><a:solidFill>{clr(line)}</a:solidFill>'
                     + (f'<a:prstDash val="{dash}"/>' if dash else "") + "</a:ln>")
        else:
            linex = "<a:ln><a:noFill/></a:ln>"
        hl = ""
        if link:
            rid = f"rIdX{self.tag}{len(self.links) + 1}"
            self.links.append((rid, link))
            hl = f'<a:hlinkClick xmlns:r="{R_NS}" r:id="{rid}"/>'
        base = dict(size=size, color=color, bold=bold, font=font)
        tx = self._txbody(text, base, align, anchor, ins, wrap) if text is not None else ""
        rotx = f' rot="{int(rot * 60000)}"' if rot else ""
        txbox = ' txBox="1"' if (fill is None and line is None and text is not None) else ""
        self._add(
            f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{self._ph(n)}" name="{esc(name)}">{hl}</xdr:cNvPr>'
            f'<xdr:cNvSpPr{txbox}/></xdr:nvSpPr>'
            f'<xdr:spPr><a:xfrm{rotx}><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'
            f'<a:prstGeom prst="{prst}"><a:avLst>{gds}</a:avLst></a:prstGeom>{fillx}{linex}</xdr:spPr>{tx}</xdr:sp>',
            x, y, w, h)
        return Shape(n, prst, x, y, w, h, adj)

    def text(self, x, y, w, h, text, size=11, color="222222", bold=False, font="BIZ UDGothic",
             align="l", anchor="t", ins=(0, 0, 0, 0), wrap="square", link=None, underline=False, name="文字"):
        if underline and isinstance(text, str):
            text = [[(p, {"underline": True})] for p in text.split("\n")]
        return self.shape(x, y, w, h, "rect", None, None, text=text, size=size, color=color, bold=bold,
                          font=font, align=align, anchor=anchor, ins=ins, wrap=wrap, link=link, name=name)

    def connector(self, start, end, color="E97132", lw=0.75, dash=None, head=False, tail=True,
                  name="矢印"):
        """真っすぐな矢印。start / end は (Shape, つなぎ目番号) か (x, y)。
        箱につなぐ端は、つなぎ目の位置から引く（Excel が後で動かしても形が変わらないように）。
        折れ線が要るときは、途中で止めた 2 本に分けるか、箱の位置をそろえて真っすぐにする。"""
        def resolve(p):
            if isinstance(p[0], Shape):
                s, idx = p
                return s.site(idx), (s.local_id, idx)
            return (p[0], p[1]), None
        (x1, y1), sc = resolve(start)
        (x2, y2), ec = resolve(end)
        x, y = min(x1, x2), min(y1, y2)
        w, h = abs(x2 - x1), abs(y2 - y1)
        flip = (' flipH="1"' if x2 < x1 else "") + (' flipV="1"' if y2 < y1 else "")
        ends = ('<a:headEnd type="triangle"/>' if head else "") + ('<a:tailEnd type="triangle"/>' if tail else "")
        dashx = f'<a:prstDash val="{dash}"/>' if dash else ""
        cx = ""
        if sc:
            cx += f'<a:stCxn id="{self._ph(sc[0])}" idx="{sc[1]}"/>'
        if ec:
            cx += f'<a:endCxn id="{self._ph(ec[0])}" idx="{ec[1]}"/>'
        n = self._new_id()
        self._add(
            f'<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="{self._ph(n)}" name="{esc(name)}"/>'
            f'<xdr:cNvCxnSpPr>{cx}</xdr:cNvCxnSpPr></xdr:nvCxnSpPr>'
            f'<xdr:spPr><a:xfrm{flip}><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'
            f'<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>'
            f'<a:ln w="{emu(lw)}"><a:solidFill>{clr(color)}</a:solidFill>{dashx}{ends}</a:ln></xdr:spPr></xdr:cxnSp>',
            x, y, w, h)

    def callout(self, x, y, w, h, text, to, fill="E8ED44", line=("accent1", None, None), color="FF0000",
                size=11, font="BIZ UDGothic", ins=(6, 4, 6, 4), name="吹き出し"):
        """線付きの吹き出し（borderCallout1）。to=(px, py) は線の先（シート上の pt）。
        線は、吹き出しの辺のうち to に向いた辺の真ん中から出す。"""
        px, py = to
        dx, dy = px - (x + w / 2), py - (y + h / 2)
        if abs(dy) * w >= abs(dx) * h:
            sy, sx = (0 if dy < 0 else 1), 0.5
        else:
            sy, sx = 0.5, (0 if dx < 0 else 1)
        adjs = {"adj1": sy * 1e5, "adj2": sx * 1e5, "adj3": (py - y) / h * 1e5, "adj4": (px - x) / w * 1e5}
        return self.shape(x, y, w, h, "borderCallout1", fill=fill, line=line, lw=1, adjs=adjs, text=text,
                          size=size, color=color, font=font, align="l", anchor="t", ins=ins, name=name)

    def cross(self, cx, cy, size=50, color=("tx1", 50, 50), name="使わない印"):
        """✕ 印（細い十字形を約 313 度回した物。人が Excel で描いた ✕ と同じ指定）。cx, cy は中心。"""
        return self.shape(cx - size / 2, cy - size / 2, size, size, "plus", fill=color, line=color,
                          adj=45833, rot=312.8, name=name)

    def plus(self, cx, cy, size=40, color=("tx1", 50, 50), name="足す印"):
        """＋ 印（二つを合わせて使う、の意味）。"""
        return self.shape(cx - size / 2, cy - size / 2, size, size, "plus", fill=color, line=color,
                          adj=40000, name=name)

    # ---- 書き出し ---------------------------------------------------------
    @staticmethod
    def _bbox(g):
        boxes = []
        for it in g.items:
            if it[0] == "leaf":
                boxes.append(it[2:6])
            else:
                boxes.append(Drawing._bbox(it[1]))
        x0 = min(b[0] for b in boxes)
        y0 = min(b[1] for b in boxes)
        x1 = max(b[0] + b[2] for b in boxes)
        y1 = max(b[1] + b[3] for b in boxes)
        return x0, y0, x1 - x0, y1 - y0

    def _group_xml(self, g, gid, top=False):
        x, y, w, h = self._bbox(g)
        xf = (f'<a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/>'
              f'<a:chOff x="{emu(x)}" y="{emu(y)}"/><a:chExt cx="{emu(w)}" cy="{emu(h)}"/>')
        descr = f' descr="xdraw:{esc(self.tag)}"' if top else ""
        body = []
        for it in g.items:
            if it[0] == "leaf":
                body.append(it[1])
            else:
                self._gcount += 1
                body.append(self._group_xml(it[1], self._ph(f"G{self._gcount}")))
        return (f'<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="{gid}" name="{esc(g.name)}"{descr}/>'
                f'<xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm>{xf}</a:xfrm></xdr:grpSpPr>'
                + "".join(body) + "</xdr:grpSp>")

    def anchor_xml(self, start_id, grid=None):
        """本当の id を振って、まとまり 1 つの XML を返す。
        grid（CellGrid）があればセルに留める（twoCellAnchor）。無ければ位置で留める（absoluteAnchor）。"""
        self._gcount = 0
        x, y, w, h = self._bbox(self.root)
        xml = self._group_xml(self.root, self._ph("G0"), top=True)
        mapping = {self._ph(i): str(start_id + i) for i in range(1, self.count + 1)}
        for k in range(self._gcount + 1):
            mapping[self._ph(f"G{k}")] = str(start_id + self.count + 1 + k)
        xml = re.sub(r"@ID[G0-9]+@", lambda m: mapping[m.group(0)], xml)
        if grid is None:
            return (f'<xdr:absoluteAnchor><xdr:pos x="{emu(x)}" y="{emu(y)}"/><xdr:ext cx="{emu(w)}" cy="{emu(h)}"/>'
                    + xml + "<xdr:clientData/></xdr:absoluteAnchor>")
        (c0, co0), (r0, ro0) = grid.col_at(emu(x)), grid.row_at(emu(y))
        (c1, co1), (r1, ro1) = grid.col_at(emu(x + w)), grid.row_at(emu(y + h))
        return (f'<xdr:twoCellAnchor><xdr:from><xdr:col>{c0}</xdr:col><xdr:colOff>{co0}</xdr:colOff>'
                f'<xdr:row>{r0}</xdr:row><xdr:rowOff>{ro0}</xdr:rowOff></xdr:from>'
                f'<xdr:to><xdr:col>{c1}</xdr:col><xdr:colOff>{co1}</xdr:colOff>'
                f'<xdr:row>{r1}</xdr:row><xdr:rowOff>{ro1}</xdr:rowOff></xdr:to>'
                + xml + "<xdr:clientData/></xdr:twoCellAnchor>")


class CellGrid:
    """セルの位置（EMU）と、列・行の番号の対応。
    Excel が書いた既存の twoCellAnchor（from/to と xfrm の両方を持つ）から、Excel の目で見た
    標準の列の幅と行の高さを割り出す。Mac の Excel では、行の高さ 20pt が 20.45pt として
    留められているなど、設定の値と実際の値がずれるため、計算式より実物から測る方が確か。
    幅や高さを個別に変えた列・行は、標準との比で伸び縮みさせる。"""

    def __init__(self, col_w, row_h, custom_cols=None, custom_rows=None):
        self.col_w, self.row_h = col_w, row_h
        self.custom_cols = custom_cols or {}   # 列番号(0始まり) -> EMU
        self.custom_rows = custom_rows or {}   # 行番号(0始まり) -> EMU

    def _at(self, pos, default, custom):
        i, acc = 0, 0
        while True:
            size = custom.get(i, default)
            if acc + size > pos:
                return i, int(pos - acc)
            acc += size
            i += 1

    def col_at(self, x):
        return self._at(x, self.col_w, self.custom_cols)

    def row_at(self, y):
        return self._at(y, self.row_h, self.custom_rows)

    @classmethod
    def measure(cls, drawing_xml, sheet_xml):
        """既存の図と sheet から測る。測れなければ None（そのときは位置で留める）。"""
        samples = []
        for a in re.findall(r"<xdr:twoCellAnchor\b.*?</xdr:twoCellAnchor>", drawing_xml, re.S):
            f = re.search(r"<xdr:from><xdr:col>(\d+)</xdr:col><xdr:colOff>(-?\d+)</xdr:colOff>"
                          r"<xdr:row>(\d+)</xdr:row><xdr:rowOff>(-?\d+)</xdr:rowOff>", a)
            t = re.search(r"<xdr:to><xdr:col>(\d+)</xdr:col><xdr:colOff>(-?\d+)</xdr:colOff>"
                          r"<xdr:row>(\d+)</xdr:row><xdr:rowOff>(-?\d+)</xdr:rowOff>", a)
            g = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/>\s*<a:ext cx="(\d+)" cy="(\d+)"', a)
            if not (f and t and g):
                continue
            tc, tco, tr, tro = (int(v) for v in t.groups())
            ox, oy, cx, cy = (int(v) for v in g.groups())
            if tc and tr:
                samples.append(((ox + cx - tco) / tc, (oy + cy - tro) / tr, tc * tr))
        if not samples:
            return None
        samples.sort(key=lambda s: -s[2])           # 一番大きい図（行・列を多くまたぐ物）を信じる
        col_w, row_h = samples[0][0], samples[0][1]
        fmt = re.search(r"<sheetFormatPr\b[^>]*>", sheet_xml)
        dcw = float(re.search(r'defaultColWidth="([\d.]+)"', fmt.group(0)).group(1)) if fmt and "defaultColWidth" in fmt.group(0) else None
        drh = float(re.search(r'defaultRowHeight="([\d.]+)"', fmt.group(0)).group(1)) if fmt else 15.0
        custom_cols = {}
        for c in re.findall(r"<col\b[^>]*>", sheet_xml):
            mn, mx = int(re.search(r'min="(\d+)"', c).group(1)), int(re.search(r'max="(\d+)"', c).group(1))
            wv = float(re.search(r'width="([\d.]+)"', c).group(1))
            if dcw and abs(wv - dcw) > 1e-6:
                for i in range(mn - 1, min(mx, 16384)):
                    custom_cols[i] = col_w * wv / dcw
        custom_rows = {}
        for r in re.findall(r"<row\b[^>]*>", sheet_xml):
            ht = re.search(r'\bht="([\d.]+)"', r)
            if ht and 'customHeight="1"' in r:
                custom_rows[int(re.search(r'\br="(\d+)"', r).group(1)) - 1] = row_h * float(ht.group(1)) / drh
        return cls(col_w, row_h, custom_cols, custom_rows)


def _strip_tagged(drawing_xml, tag):
    """前に同じ目印で足したまとまりを消す。"""
    out, pos, n = [], 0, 0
    pat = re.compile(r"<xdr:(absoluteAnchor|oneCellAnchor|twoCellAnchor)\b.*?</xdr:\1>", re.S)
    for m in pat.finditer(drawing_xml):
        if f'descr="xdraw:{tag}"' in m.group(0):
            out.append(drawing_xml[pos:m.start()])
            pos = m.end()
            n += 1
    out.append(drawing_xml[pos:])
    return "".join(out), n


def _rels_with_links(rels_xml, tag, links):
    if rels_xml is None:
        rels_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
    rels_xml = re.sub(rf'<Relationship Id="rIdX{re.escape(tag)}\d+"[^>]*/>', "", rels_xml)
    add = "".join(f'<Relationship Id="{rid}" Type="{HLINK_TYPE}" Target="{html.escape(url, quote=True)}" '
                  f'TargetMode="External"/>' for rid, url in links)
    return rels_xml.replace("</Relationships>", add + "</Relationships>")


def _sheet_of_drawing(zin, drawing_part):
    """この drawing を使っている sheet の部品名。"""
    base = drawing_part.rsplit("/", 1)[1]
    for n in zin.namelist():
        if n.startswith("xl/worksheets/_rels/") and n.endswith(".rels"):
            if re.search(rf'Target="[^"]*drawings/{re.escape(base)}"', zin.read(n).decode("utf-8")):
                return "xl/worksheets/" + n.rsplit("/", 1)[1][:-5]
    return None


def measure_grid(xlsx, drawing_part="xl/drawings/drawing1.xml"):
    """xlsx の既存の図から、セルの大きさ（CellGrid）を測る。人の図を外して描き直す前に呼ぶ。"""
    with zipfile.ZipFile(xlsx) as z:
        sheet = _sheet_of_drawing(z, drawing_part)
        if not sheet:
            return None
        return CellGrid.measure(z.read(drawing_part).decode("utf-8"), z.read(sheet).decode("utf-8"))


def _strip_untagged(drawing_xml):
    """台本で足していない図（人が描いた図）をすべて外す。描き直すときだけ使う。"""
    pat = re.compile(r"<xdr:(absoluteAnchor|oneCellAnchor|twoCellAnchor)\b.*?</xdr:\1>", re.S)
    n = 0

    def drop(m):
        nonlocal n
        if 'descr="xdraw:' in m.group(0):
            return m.group(0)
        n += 1
        return ""
    return pat.sub(drop, drawing_xml), n


def write_into(xlsx_in, xlsx_out, drawings, drawing_part="xl/drawings/drawing1.xml", anchor="cell",
               grid=None, replace_hand_drawn=False, sheet_edit=None):
    """xlsx_in の図（drawing_part）に drawings（Drawing 1 つか list）のまとまりを足して xlsx_out に書く。
    anchor="cell" はセルに留める（人が描いた図と一緒に伸び縮みする。既定）。grid を渡せばその大きさで留める。
    セルの大きさを測れないとき、または anchor="absolute" のときは位置で留める。
    同じ目印のまとまり・リンクが既にあれば消してから足す。
    replace_hand_drawn=True のときは、人が描いた図（目印の無い図）も外す。先に measure_grid で grid を測っておくこと。
    sheet_edit は sheet の XML を受け取って直した XML を返す関数（セルの文字を動かすときなど）。
    ほかの部品は変えない。戻り値は外した図の数。"""
    if not isinstance(drawings, (list, tuple)):
        drawings = [drawings]
    rels_part = drawing_part.replace("drawings/", "drawings/_rels/") + ".rels"
    with zipfile.ZipFile(xlsx_in) as zin:
        names = zin.namelist()
        if drawing_part not in names:
            raise SystemExit(f"{drawing_part} がありません。先に Excel で図形を 1 つ置いてから使ってください。")
        sheet = _sheet_of_drawing(zin, drawing_part)
        dxml = zin.read(drawing_part).decode("utf-8")
        if anchor == "cell" and grid is None and sheet:
            grid = CellGrid.measure(dxml, zin.read(sheet).decode("utf-8"))
        removed = 0
        if replace_hand_drawn:
            dxml, removed = _strip_untagged(dxml)
        for d in drawings:
            dxml, n = _strip_tagged(dxml, d.tag)
            removed += n
        if not dxml.rstrip().endswith("</xdr:wsDr>"):
            raise SystemExit("drawing の終わりが想定と違う")
        rels = zin.read(rels_part).decode("utf-8") if rels_part in names else None
        for d in drawings:
            used = [int(v) for v in re.findall(r'<xdr:cNvPr id="(\d+)"', dxml)] or [0]
            dxml = dxml[: dxml.rfind("</xdr:wsDr>")] + d.anchor_xml(max(used) + 1, grid if anchor == "cell" else None) + "</xdr:wsDr>"
            rels = _rels_with_links(rels, d.tag, d.links)
        with zipfile.ZipFile(xlsx_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == drawing_part:
                    data = dxml.encode("utf-8")
                elif info.filename == rels_part:
                    data = rels.encode("utf-8")
                elif sheet_edit and info.filename == sheet:
                    data = sheet_edit(data.decode("utf-8")).encode("utf-8")
                zout.writestr(info, data)
            if rels_part not in names:
                zout.writestr(rels_part, rels.encode("utf-8"))
    return removed
