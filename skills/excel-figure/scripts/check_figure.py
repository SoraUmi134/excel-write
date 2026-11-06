#!/usr/bin/env python3
"""Excel の説明図を機械的に点検する。人に渡す前に必ず通す（渡す前の点検の skill から呼ぶ）。

使い方:
  python3 check_figure.py 図.xlsx [--drawing xl/drawings/drawing1.xml] [--group パターン3-B]

見つける物（人に指摘されたことを、機械で拾える形にした物）:
  あふれ      文が箱からあふれる（Excel の字の広さで見積もる。見積もりは Excel より少し厳しめ）
  凡例        凡例に無い線の種類が使われている／凡例にあるのに使っていない
  向き        「読む」線が、読まれる物（計画書・手引き）から出ている
  横切り      線や吹き出しの線が、ほかの箱を横切っている
  重なり      文を持つ図形どうしが重なっている
  はみ出し    枠の線をまたいでいる（吹き出しとその札は除く）
  禁止語      使ってはいけない言い方（../../review-before-delivery/banned_words.txt）
  長い文      1 つの図形の文が長すぎる

終了コード: 見つかった数が 0 なら 0、あれば 1。
"""
import argparse
import html
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import render_drawing as R  # noqa: E402

EMU = 12700
# Excel（Mac）で実物を見て測った字体ごとの値（2026-11-06）。(漢字・かなの幅, 英数字の幅, 1 行の高さ) × 字の大きさ
FONT_METRICS = {"BIZ UDGothic": (1.0, 0.5, 1.32), "Yu Gothic": (1.03, 0.55, 1.5)}
DEFAULT_FONT = "Yu Gothic"
SLACK = 1.02           # 見積もりの誤差の分だけ、幅に余裕を見る
LONG_PARA = 75         # 1 段落の長さの上限（字）。吹き出しは LONG_PARA_CALLOUT
LONG_PARA_CALLOUT = 95
# 線の意味の名前（台本の部品が付ける名前）と、その意味の線が同じ見た目であるべき組
LINE_KINDS = {"コマンド": "cmd", "起こす矢印": "spawn", "起こす矢印の幹": "spawn", "起こす矢印の枝": "spawn",
              "書く": "write", "読む": "read", "できない線": "cannot", "区切り線": "separator"}
HERE = os.path.dirname(os.path.abspath(__file__))
BANNED = os.path.join(HERE, "..", "..", "review-before-delivery", "banned_words.txt")
DOC_WORDS = ("計画書", "PLAN.md", "SKILL.md", "手引き", "capability.json")
IGNORE_OVERLAP = ("重なり", "できない所", "使わない印", "足す印", "番号", "凡例", "区切り線")


class Shape:
    def __init__(self, **kw):
        self.__dict__.update(kw)

    @property
    def rect(self):
        return (self.x, self.y, self.x + self.w, self.y + self.h)


def parse(xlsx, drawing):
    z = zipfile.ZipFile(xlsx)
    R.load_theme(z)
    xml = z.read(drawing).decode("utf-8")
    shapes = []

    def walk(part_xml, xf, groups):
        for tag, part in R.split_children(part_xml):
            if tag == "xdr:grpSp":
                gx = R.first(r"<xdr:grpSpPr>\s*(<a:xfrm.*?</a:xfrm>)", part)
                gname = html.unescape(R.first(r'<xdr:nvGrpSpPr><xdr:cNvPr[^>]*name="([^"]*)"', part) or "")
                inner = part[part.find("</xdr:grpSpPr>") + len("</xdr:grpSpPr>"):part.rfind("</xdr:grpSp>")]
                walk(inner, xf.child(gx) if gx else xf, groups + [gname])
            elif tag in ("xdr:sp", "xdr:cxnSp"):
                shapes.append(one(part, xf, groups, tag == "xdr:cxnSp"))

    for m in re.finditer(r"<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b.*?</xdr:\1>", xml, re.S):
        walk(m.group(0), R.Xform(), [])
    return [s for s in shapes if s]


def one(part, xf, groups, is_cxn):
    sppr = R.first(r"(<xdr:spPr\b.*?</xdr:spPr>)", part) or ""
    off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/>\s*<a:ext cx="(\d+)" cy="(\d+)"', sppr)
    if not off:
        return None
    x, y = xf.pt(int(off.group(1)), int(off.group(2)))
    w, h = xf.scale(int(off.group(3)), int(off.group(4)))
    flipH = bool(re.search(r'<a:xfrm[^>]*flipH="1"', sppr))
    flipV = bool(re.search(r'<a:xfrm[^>]*flipV="1"', sppr))
    prst = R.first(r'prst="(\w+)"', sppr) or "rect"
    sid = R.first(r'<xdr:cNvPr id="(\d+)"', part)
    name = html.unescape(R.first(r'<xdr:cNvPr[^>]*name="([^"]*)"', part) or "")
    ls = R.line_style(sppr, R.first(r"(<xdr:style>.*?</xdr:style>)", part) or "")
    line = (ls[0], bool(ls[2])) if ls else None
    tb = R.first(r"<xdr:txBody>(.*?)</xdr:txBody>", part) or ""
    body = R.first(r"(<a:bodyPr\b[^>]*?(?:/>|>.*?</a:bodyPr>))", tb) or ""
    ins = [R.first(rf'{k}Ins="(\d+)"', body) for k in "ltrb"]
    ins = [int(v) / EMU if v else d for v, d in zip(ins, (7.2, 3.6, 7.2, 3.6))]
    paras = []
    for pm in re.finditer(r"<a:p>(.*?)</a:p>|<a:p/>", tb, re.S):
        pa = pm.group(1) or ""
        runs = []
        for r in re.findall(r"<a:r>.*?</a:r>", pa, re.S):
            sz = int(R.first(r'sz="(\d+)"', r) or 1100) / 100
            b = R.first(r'\bb="(\d)"', r) == "1"
            font = html.unescape(R.first(r'<a:latin typeface="([^"]+)"', r) or DEFAULT_FONT)
            runs.append((html.unescape("".join(re.findall(r"<a:t>(.*?)</a:t>", r, re.S))), sz, (b, font)))
        if not runs:
            runs = [("", int(R.first(r'sz="(\d+)"', pa) or 1100) / 100, (False, DEFAULT_FONT))]
        paras.append(runs)
    text = "".join(t for p in paras for t, _, _ in p)
    adjs = {k: int(v) / 1e5 for k, v in re.findall(r'<a:gd name="(adj\d)" fmla="val (-?\d+)"', sppr)}
    return Shape(id=sid, name=name, prst=prst, x=x, y=y, w=w, h=h, flipH=flipH, flipV=flipV,
                 is_cxn=is_cxn or prst in ("line", "straightConnector1") or prst.startswith("bentConnector"),
                 line=line, ins=ins, paras=paras, text=text, wrap=R.first(r'wrap="(\w+)"', body) or "square",
                 groups=groups, adjs=adjs,
                 st=R.first(r'<a:stCxn id="(\d+)"', part), en=R.first(r'<a:endCxn id="(\d+)"', part),
                 has_fill=not re.search(r"</a:prstGeom>\s*<a:noFill/>", sppr))


def metrics(font):
    return FONT_METRICS.get(font, FONT_METRICS[DEFAULT_FONT])


def char_w(ch, sz, style):
    bold, font = style
    wide, narrow, _ = metrics(font)
    o = ord(ch)
    if o < 0x7F:
        w = sz * narrow
    elif 0xFF61 <= o <= 0xFF9F:
        w = sz * 0.5
    else:
        w = sz * wide
    return w * (1.05 if bold else 1.0)


def need_height(s):
    """折り返したときに要る高さ（pt）。"""
    maxw = (s.w - s.ins[0] - s.ins[2]) * SLACK
    total = 0.0
    for runs in s.paras:
        lines, cur = 1, 0.0
        line_h = max(sz * metrics(st[1])[2] for _, sz, st in runs)
        for t, sz, st in runs:
            for ch in t:
                cw = char_w(ch, sz, st)
                if s.wrap != "none" and cur + cw > maxw and cur > 0:
                    lines += 1
                    cur = 0.0
                cur += cw
        total += lines * line_h
    return total + s.ins[1] + s.ins[3]


def need_width(s):
    return max(sum(char_w(ch, sz, st) for t, sz, st in runs for ch in t) for runs in s.paras) + s.ins[0] + s.ins[2]


def text_rect(s):
    """文が実際に占める範囲（折り返さない文は、文の長さまで）。"""
    if s.wrap == "none":
        return (s.x, s.y, s.x + min(s.w, need_width(s)), s.y + s.h)
    return s.rect


def seg_hits_rect(x1, y1, x2, y2, r, pad=2):
    """線分が四角の内側を通るか（端の近くは除く）。"""
    rx0, ry0, rx1, ry1 = r[0] + pad, r[1] + pad, r[2] - pad, r[3] - pad
    if rx0 >= rx1 or ry0 >= ry1:
        return False
    steps = 40
    for k in range(3, steps - 2):
        t = k / steps
        px, py = x1 + (x2 - x1) * t, y1 + (y2 - y1) * t
        if rx0 < px < rx1 and ry0 < py < ry1:
            return True
    return False


def inside(a, b, pad=1):
    return a[0] >= b[0] - pad and a[1] >= b[1] - pad and a[2] <= b[2] + pad and a[3] <= b[3] + pad


def overlap_ratio(a, b):
    ix = max(0, min(a[2], b[2]) - max(a[0], b[0]))
    iy = max(0, min(a[3], b[3]) - max(a[1], b[1]))
    small = min((a[2] - a[0]) * (a[3] - a[1]), (b[2] - b[0]) * (b[3] - b[1])) or 1
    return ix * iy / small


def load_banned():
    words = []
    if os.path.exists(BANNED):
        for line in open(BANNED, encoding="utf-8"):
            line = line.split("#", 1)[0].strip()
            if line:
                words.append(line.split("\t")[0].strip())
    return words


def check(shapes, only_group=None):
    found = []

    def add(kind, s, msg):
        top = s.groups[0] if s.groups else "（まとまりの外）"
        found.append((top, kind, s.name, msg))

    if only_group:
        shapes = [s for s in shapes if s.groups and s.groups[0] == only_group]
    by_id = {s.id: s for s in shapes}
    tops = sorted({s.groups[0] for s in shapes if s.groups})
    banned = load_banned()
    boxes = [s for s in shapes if not s.is_cxn and s.has_fill and s.text.strip() and s.w > 30 and s.h > 20
             and not any(k in s.name for k in IGNORE_OVERLAP)]
    for s in shapes:
        if s.is_cxn or not s.text.strip():
            continue
        if s.wrap != "none":
            need = need_height(s)
            if need > s.h + 3:
                add("あふれ", s, f"高さ {s.h:.0f}pt に対し {need:.0f}pt 要る: 「{s.text[:30]}…」")
        for wd in banned:
            if wd in s.text:
                add("禁止語", s, f"「{wd}」を使っている: 「{s.text[:40]}…」")
        limit = LONG_PARA_CALLOUT if s.prst.startswith("borderCallout") else LONG_PARA
        for runs in ([] if "人の文" in s.name else s.paras):
            para = "".join(t for t, _, _ in runs)
            if len(para) > limit:
                add("長い文", s, f"1 段落 {len(para)} 字（上限 {limit}）: 「{para[:30]}…」")
    for top in tops:
        grp = [s for s in shapes if s.groups and s.groups[0] == top]
        legend = {s.line for s in grp if s.is_cxn and s.line and "凡例" in "".join(s.groups[1:]) + s.name}
        used = {}
        for s in grp:
            if s.is_cxn and s.line and "凡例" not in "".join(s.groups[1:]) + s.name and s.name != "区切り線":
                used.setdefault(s.line, s.name)
        kinds = {}
        for s in grp:
            if not s.is_cxn or not s.line or "凡例" in "".join(s.groups[1:]) + s.name:
                continue
            kind = LINE_KINDS.get(s.name.replace("（横切り可）", ""))
            if kind is None:
                add("線の意味", s, "線に意味の名前が無い（コマンド・起こす矢印・書く・読む など。台本の部品で引く）")
                continue
            kinds.setdefault(kind, set()).add(s.line)
        for kind, styles in kinds.items():
            if len(styles) > 1 and kind != "cannot":
                found.append((top, "線の意味", kind, f"同じ意味の線なのに見た目が {len(styles)} 通りある"))
        for k1 in kinds:
            for k2 in kinds:
                if k1 < k2 and kinds[k1] & kinds[k2] and "separator" not in (k1, k2):
                    found.append((top, "線の意味", f"{k1}/{k2}", "違う意味の線が同じ見た目（凡例で見分けられない）"))
        if legend:
            for st, nm in used.items():
                if st not in legend:
                    found.append((top, "凡例", nm, f"凡例に無い線の種類（色 {st[0]}、{'点線' if st[1] else '実線'}）"))
            for st in legend - set(used):
                found.append((top, "凡例", "凡例", f"使っていない凡例（色 {st[0]}、{'点線' if st[1] else '実線'}）"))
        elif used:
            found.append((top, "凡例", "（なし）", f"線が {len(used)} 種類あるのに凡例が無い"))
        for s in grp:
            if s.is_cxn and s.name == "読む" and s.st in by_id:
                src = by_id[s.st]
                if any(k in src.text for k in DOC_WORDS) and "サブエージェント" not in src.text:
                    add("向き", s, f"読む線が「{src.text[:20]}…」から出ている（読む側から読まれる物へ向ける）")
        frame = next((s for s in grp if s.name == "枠"), None)
        if frame:
            fr = frame.rect
            for s in grp:
                if s is frame or s.is_cxn or s.prst.startswith("borderCallout"):
                    continue
                if any(g.startswith("問題点") for g in s.groups) or s.name == "見出し":
                    continue
                r = text_rect(s)
                crosses = (r[0] < fr[0] < r[2] or r[0] < fr[2] < r[2] or r[1] < fr[1] < r[3] or r[1] < fr[3] < r[3])
                if crosses and s.text.strip():
                    add("はみ出し", s, f"枠の線をまたいでいる: 「{s.text[:20]}…」")
    for s in shapes:
        if not s.is_cxn or not s.line or s.name in ("区切り線",) or "凡例" in s.name or "横切り可" in s.name:
            continue
        x1, x2 = (s.x + s.w, s.x) if s.flipH else (s.x, s.x + s.w)
        y1, y2 = (s.y + s.h, s.y) if s.flipV else (s.y, s.y + s.h)
        for b in boxes:
            if b.id in (s.st, s.en) or (b.groups[:1] != s.groups[:1]):
                continue
            if seg_hits_rect(x1, y1, x2, y2, b.rect):
                add("横切り", s, f"線が「{b.text[:16]}…」の箱を横切っている")
    for c in shapes:
        if not c.prst.startswith("borderCallout"):
            continue
        a = c.adjs
        sx, sy = c.x + c.w * a.get("adj2", 0), c.y + c.h * a.get("adj1", 0)
        ex, ey = c.x + c.w * a.get("adj4", 0), c.y + c.h * a.get("adj3", 0)
        for b in boxes:
            if b is c or b.groups[:1] != c.groups[:1]:
                continue
            r = b.rect
            if r[0] - 3 <= ex <= r[2] + 3 and r[1] - 3 <= ey <= r[3] + 3:
                continue
            if seg_hits_rect(sx, sy, ex, ey, r):
                add("横切り", c, f"吹き出しの線が「{b.text[:16]}…」の箱を横切っている")
    texts = [s for s in shapes if not s.is_cxn and s.text.strip() and not any(k in s.name for k in IGNORE_OVERLAP)
             and not s.prst.startswith("borderCallout") and not s.name.startswith("参考リンク") and "重なり可" not in s.name]
    for i, a in enumerate(texts):
        for b in texts[i + 1:]:
            if a.groups[:1] != b.groups[:1]:
                continue
            ra, rb = text_rect(a), text_rect(b)
            if inside(ra, rb) or inside(rb, ra):
                continue
            if overlap_ratio(ra, rb) > 0.12:
                add("重なり", a, f"「{a.text[:14]}…」と「{b.text[:14]}…」が重なっている")
    return found


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("xlsx")
    ap.add_argument("--drawing", default="xl/drawings/drawing1.xml")
    ap.add_argument("--group", help="このまとまり（パターン）だけを見る")
    a = ap.parse_args()
    found = check(parse(a.xlsx, a.drawing), a.group)
    if not found:
        print("点検: 問題は見つからなかった（機械で拾える範囲。伝わるかは予備知識の無い係に図の画像を見せて確かめる）")
        return 0
    print(f"点検: {len(found)} 件")
    for top, kind, name, msg in sorted(found):
        print(f"  [{kind}] {top} / {name}: {msg}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
