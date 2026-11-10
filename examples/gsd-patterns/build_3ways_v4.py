"""「GSDに独自処理を挟む3つの方法」の v4 を、台本で描き直す。

使い方:
  python3 build_3ways_v4.py v3.xlsx v4.xlsx

- v3 の中身（文・色・字体・箱の形・参考リンクの行き先）を手本に、図 1〜3 と比べの表を描き直す。
  v3 の図（目印 v2*）は外し、このファイルの図（目印 v4*）に置き換える。参考リンクは v3 にある行き先だけを使う。
- 骨組み（上の 4 つのコマンドの札・メインの帯・コマンドごとの境目・左に凡例・右に吹き出し・下に比べの表）は v3 と同じ。
- 問題点の番号は、比べの表の行の番号（1〜9）にそろえる。図の中の赤・金の丸の数字と、吹き出しの「問題点 N」と、表の行が同じ番号。
- 手順の番号は 3 図とも (1) の形。矢印の横に置き、線に重ねない。
- 線の色は 1 つの意味だけに使う（LINES）。凡例も同じ表から描く。赤は「できない」にだけ使う。
"""
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "skills", "excel-figure", "scripts"))
from xdraw import Drawing, measure_grid, write_into  # noqa: E402

# ---- 色・字（v3 から写した物） ------------------------------------------------
GREEN = ("accent6", 20, 80)        # GSD の物
GREEN_LINE = ("accent3",)
BLUE = ("accent1", 20, 80)         # 方法論（キット）の物
BLUE_LINE = "2F6DB5"
PURPLE = ("accent5",)              # 流れを曲げる設定・決まり（紫の字）
LINK = ("accent4",)
RED = "FF0000"                     # できない（✕）にだけ使う
GOLD = "BF8F00"                    # 確実でない・作り直しが要る（△）
OK_GREEN = "38761D"                # ○
INK = "222222"
SUB = "595959"
UD = "BIZ UDGothic"
YU = "Yu Gothic"
AW = 1.5
CALLOUT_FILL = "FFF7CC"            # 吹き出しの地（淡い黄）
CALLOUT_LINE = "BF9000"
FRAME_LINE = "8C8C8C"              # まとめて扱う物の点線の枠
GHOST_FILL, GHOST_LINE = "F2F2F2", "7F7F7F"

# 線の種類（凡例と実際の線を同じ表から引く）
LINES = {
    "cmd": ("404040", AW, None, "人のコマンドで、メインエージェントが動く", "コマンド"),
    "spawn": (("accent2",), AW, None, "サブエージェントを起こす（呼び出す）", "起こす矢印"),
    "write": (("accent3",), AW, None, "計画書を書く", "書く"),
    "read": (("tx1", 50, 50), 1.25, "dash", "読んで従う・読んで見張る（矢印の先の物を読む）", "読む"),
    "back": ("0E7C86", AW, None, "結果をメインへ返す", "返す"),
    "cannot": (RED, AW, "dash", "✕ 本当は要るが、この形ではできない", "できない線"),
    "boundary": (("tx1", 35, 65), 1.0, "sysDash", "コマンドごとの境目（その上のコマンドで動く範囲）", "境目"),
}

GH = "https://github.com/open-gsd/gsd-core"
L = {   # v3 にある行き先だけ
    "exec_routing": GH + "/blob/603b2c6/gsd-core/workflows/execute-phase.md#:~:text=Executor%20routing",
    "4139": GH + "/blob/main/docs/adr/4139-compact-content-seam.md#:~:text=prose%20dispatch%20is%20not%20guaranteed",
    "executor_md": GH + "/blob/603b2c6/agents/gsd-executor.md",
    "tools": GH + "/blob/main/docs/FEATURES.md#:~:text=no%20capability%20can%20grant%20tools%20to%20one",
    "ship": GH + "/blob/main/docs/how-to/develop-a-capability.md#ship-skills--and-know-what-you-are-shipping",
    "1689": GH + "/issues/1689#issuecomment-5143711828",
    "4740": GH + "/issues/4740",
    "1143": GH + "/blob/main/docs/adr/1143-claude-orchestration-capability.md#:~:text=the%20orchestrator%20IS%20the%20host",
    "step275": GH + "/blob/main/gsd-core/workflows/execute-phase.md#:~:text=Contribution%20dispatch",
    "planmd": GH + "/blob/main/docs/reference/plan-md.md#per-plan-executor-routing",
    "firstparty": GH + "/blob/main/docs/explanation/capability-overlay-model.md#first-party-always-wins",
    "2363": GH + "/blob/main/docs/adr/2363-capability-instruction-surface-trust.md#:~:text=Third%2Dparty%20agents%5B%5D%20are%20never%20staged%20into%20the%20agent%27s%20instruction%20context.",
    "manifest": GH + "/blob/main/docs/reference/capability-manifest.md#feature-body-role-feature",
    "5738": GH + "/issues/5738#issuecomment-6476728511",
    "planmd_603": GH + "/blob/603b2c6/docs/reference/plan-md.md#per-plan-executor-routing",
    "skills_603": GH + "/blob/603b2c6/docs/CONFIGURATION.md#agent-skills-injection",
    "routing": GH + "/blob/main/docs/CONFIGURATION.md#workflow-toggles:~:text=workflow.agent_hint_routing",
    "settings": "https://docs.opengsd.net/core/configuration/settings#section-breakdown",
}

# ---- 字の幅の見積もり（check_figure.py と同じ値。Excel は 1〜2 割広いので SLACK で余裕を取る） ----
METRICS = {UD: (1.0, 0.5, 1.32), YU: (1.03, 0.55, 1.5), None: (1.03, 0.55, 1.5)}
SLACK = 1.15


def _norm(text, base):
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


def need_h(text, w, size=10, font=UD, ins=(6, 4, 6, 4)):
    """折り返したときに要る高さ（pt）。"""
    maxw = w - ins[0] - ins[2]
    total = 0.0
    for runs in _norm(text, {"size": size, "font": font}):
        lines, cur = 1, 0.0
        lh = max(o["size"] * METRICS.get(o["font"], METRICS[None])[2] for _, o in runs)
        for t, o in runs:
            wide, narrow, _ = METRICS.get(o["font"], METRICS[None])
            for ch in t:
                cw = o["size"] * (narrow if ord(ch) < 0x7F else wide) * (1.05 if o.get("bold") else 1) * SLACK
                if cur + cw > maxw and cur > 0:
                    lines += 1
                    cur = 0.0
                cur += cw
        total += lines * lh
    return total + ins[1] + ins[3]


# ---- 部品 -----------------------------------------------------------------
def box(d, x, y, w, h, lines, fill=GREEN, line=GREEN_LINE, size=10, name="箱", align="ctr", prst="roundRect",
        dash=None, anchor="ctr"):
    return d.shape(x, y, w, h, prst, fill=fill, line=line, lw=0.75, dash=dash,
                   adj=8000 if prst == "roundRect" else None, align=align, anchor=anchor,
                   text=lines, size=size, font=YU, color=INK, ins=(6 if align == "l" else 4, 2, 4, 2), name=name)


def doc(d, x, y, w, h, lines, size=9, name="書類", align="ctr"):
    """読むだけの物（設定・手引き・依頼文に書き足す文）。紙の書類の形で、縁は点線。"""
    return box(d, x, y, w, h, lines, fill=BLUE, line=BLUE_LINE, size=size, name=name, align=align,
               prst="flowChartDocument", dash="dash", anchor="t")


def stack(d, x, y, w, h, lines, fill=GREEN, line=GREEN_LINE, size=10, name="重なった箱"):
    """同じ物が画面の数だけあることを示す、重なったカード。前の 1 枚を返す。"""
    with d.group(name):
        for k in (2, 1):
            d.shape(x + 4 * k, y + 3 * k, w, h, "roundRect", fill=fill, line=line, lw=0.75, adj=8000, name="重なり")
        return box(d, x, y, w, h, lines, fill, line, size, name)


def frame(d, x, y, w, h, name="係の枠"):
    """まとめて扱う物を囲む点線の枠。線は枠に刺す。"""
    return d.shape(x, y, w, h, "rect", fill=None, line=FRAME_LINE, lw=1, dash="dash", name=name)


def ghost(d, x, y, w, h, lines, name="本当は要る物"):
    return d.shape(x, y, w, h, "roundRect", fill=GHOST_FILL, line=GHOST_LINE, lw=1, dash="dash", adj=8000,
                   text=lines, size=10, font=YU, color=SUB, ins=(4, 2, 4, 2), name=name)


def auto_check(d, x, y, w, h, lines, name="自動の検査"):
    """自動で動く検査（hook・Capability の門）。両端に帯のある箱で、サブエージェントと分ける。"""
    return d.shape(x, y, w, h, "flowChartPredefinedProcess", fill=BLUE, line=BLUE_LINE, lw=0.75,
                   text=lines, size=10, font=YU, color=INK, ins=(16, 2, 16, 2), name=name)


def line(d, kind, a, b, tail=True):
    color, lw, dash, _, name = LINES[kind]
    d.connector(a, b, color=color, lw=lw, dash=dash, tail=tail, name=name)


def number(d, x, y, label):
    """手順の番号（(1) の形）。x, y は左上。矢印の横に置き、線に重ねない。"""
    d.text(x, y, 48, 26, label, size=18, font=None, color="000000", align="l", wrap="none", name="番号")


def badge(d, cx, cy, n, gold=False, size=20):
    """丸の数字。比べの表の行の番号と同じ。赤＝できない、金＝確実でない・作り直しが要る。"""
    return d.shape(cx - size / 2, cy - size / 2, size, size, "ellipse", fill=GOLD if gold else RED, line=None,
                   text=[[(str(n), {"bold": True, "color": "FFFFFF"})]], size=11, font=UD,
                   ins=(0, 0, 0, 0), name=f"できない所 {n}")


def cross(d, cx, cy, size, color):
    d.cross(cx, cy, size=size, color=color)


def link_tags(d, x, y, keys, per_row=5):
    for i, k in enumerate(keys):
        d.text(x + (i % per_row) * 82, y + (i // per_row) * 20, 78, 20, f"参考リンク{i + 1}", size=11, font=UD,
               color=LINK, underline=True, link=L[k], wrap="none", name=f"参考リンク{i + 1}")


def note(d, x, y, w, text, size=9.5, color="000000", name="注", h=None):
    h = h or need_h(text, w, size, UD, (0, 0, 0, 0)) + 2
    d.text(x, y, w, h, text, size=size, font=UD, color=color, name=name)
    return y + h


# ---- 吹き出し ------------------------------------------------------------------
CX, CW = 1580, 470            # 吹き出しの列（図の枠の外の右）


def problem_size(head, body, keys):
    text = [[(head, {"bold": True, "size": 10.5})]] + body
    h = need_h(text, CW, 10, UD, (26, 5, 8, 4))
    return text, h + (24 if keys else 4)


def problem(d, y, nums, gold, head, body, to, keys):
    """問題点の吹き出し。nums は比べの表の行の番号。見出しの左に同じ丸の数字を付ける。
    body は段落の list。段落は [(文, {"color": ...}), ...]。何が困るかの所だけ赤にする。"""
    text, h = problem_size(head, body, keys)
    with d.group("問題点 " + "・".join(map(str, nums))):
        # 線はいつも吹き出しの左の辺の真ん中から出す（右の列のほかの吹き出しを横切らないように）
        adjs = {"adj1": 50000, "adj2": 0, "adj3": (to[1] - y) / h * 1e5, "adj4": (to[0] - CX) / CW * 1e5}
        d.shape(CX, y, CW, h, "borderCallout1", fill=CALLOUT_FILL, line=CALLOUT_LINE, lw=1, adjs=adjs, text=text,
                size=10, color=INK, font=UD, align="l", anchor="t", ins=(26, 5, 8, 4), name="吹き出し")
        for k, n in enumerate(nums):
            badge(d, CX + 13, y + 13 + 22 * k, n, gold=gold[k])
        if keys:
            link_tags(d, CX + 26, y + h - 23, keys)
    return y + h


def place_problems(d, items, top, gap=8):
    """吹き出しを、話の対象の高さになるべく近く、重ならないように置く（items は流れの順）。
    重なる吹き出しは一塊にして、塊の真ん中が対象の高さの平均に来るようにずらす。
    yrange がある物は、箱の辺のどこを指してもよいので、吹き出しの高さに合わせて横に近い線にする。"""
    hs = [problem_size(it["head"], it["body"], it["keys"])[1] for it in items]
    ideal = [it["to"][1] - h / 2 for it, h in zip(items, hs)]
    blocks = [[i] for i in range(len(items))]

    def pos(b):
        offs, acc = [], 0
        for i in b:
            offs.append(acc)
            acc += hs[i] + gap
        y = sum(ideal[i] - o for i, o in zip(b, offs)) / len(b)
        return max(y, top), offs

    while True:
        merged = False
        for k in range(len(blocks) - 1):
            y0, offs = pos(blocks[k])
            y1, _ = pos(blocks[k + 1])
            end = y0 + offs[-1] + hs[blocks[k][-1]] + gap
            if end > y1:
                blocks[k:k + 2] = [blocks[k] + blocks[k + 1]]
                merged = True
                break
        if not merged:
            break
    bottom = top
    for b in blocks:
        y0, offs = pos(b)
        for i, o in zip(b, offs):
            it, y = items[i], y0 + o
            to = it["to"]
            if "yrange" in it:
                rs = it["yrange"] if isinstance(it["yrange"], list) else [it["yrange"]]
                c = y + hs[i] / 2
                to = (to[0], min((min(max(c, lo), hi) for lo, hi in rs), key=lambda v: abs(v - c)))
            bottom = problem(d, y, it["nums"], it["gold"], it["head"], it["body"], to, it["keys"])
    return bottom


def P(*runs):
    """吹き出しの段落。文字列は濃い字、("…", "red") は赤。"""
    out = []
    for r in runs:
        if isinstance(r, tuple):
            out.append((r[0], {"color": RED}))
        else:
            out.append((r, {"color": INK}))
    return out


# ---- 凡例 -----------------------------------------------------------------------
def legend(d, x, y, lines, shapes):
    """凡例。lines は LINES の鍵、shapes は見た目の鍵。図で使う物だけを載せる。"""
    with d.group("凡例"):
        d.text(x, y, 120, 18, [[("凡例", {"bold": True})]], size=10, font=UD, color="000000", wrap="none",
               name="凡例の見出し")
        yy = y + 20
        for k in lines:
            color, lw, dash, label, _ = LINES[k]
            d.connector((x, yy + 9), (x + 34, yy + 9), color=color, lw=lw, dash=dash, name="凡例の線")
            d.text(x + 40, yy, 300, 18, label, size=10, font=UD, color=RED if k == "cannot" else "000000",
                   wrap="none", name="凡例の文")
            yy += 19
        yy += 4
        for k in shapes:
            if k == "colors":
                d.shape(x, yy + 2, 16, 14, "roundRect", fill=GREEN, line=GREEN_LINE, adj=8000, name="凡例の色")
                d.shape(x + 18, yy + 2, 16, 14, "roundRect", fill=BLUE, line=BLUE_LINE, adj=8000, name="凡例の色")
                label = "緑の箱 = GSD の物　青の箱 = 方法論（キット）の物"
            elif k == "stack":
                for j in (2, 1, 0):
                    d.shape(x + 3 * j, yy + 1 + 2 * j, 26, 12, "roundRect", fill=GREEN, line=GREEN_LINE, adj=8000,
                            name="凡例の重なり")
                label = "重ねたカード = 画面の数だけある"
            elif k == "doc":
                d.shape(x, yy + 1, 34, 16, "flowChartDocument", fill=BLUE, line=BLUE_LINE, dash="dash",
                        name="凡例の書類")
                label = "書類の形 = 読むだけの物（設定・手引き・依頼文に書き足す文）"
            elif k == "frame":
                d.shape(x, yy + 1, 34, 16, "rect", fill=None, line=FRAME_LINE, lw=1, dash="dash", name="凡例の枠")
                label = "点線の枠 = まとめて扱う物（枠に刺した線は、中の全員に掛かる）"
            elif k == "auto":
                d.shape(x, yy + 1, 34, 16, "flowChartPredefinedProcess", fill=BLUE, line=BLUE_LINE,
                        name="凡例の自動の検査")
                label = "両端に帯のある箱 = 自動で動く検査（hook・Capability の門）"
            elif k == "ghost":
                d.shape(x, yy + 1, 34, 16, "roundRect", fill=GHOST_FILL, line=GHOST_LINE, dash="dash", adj=8000,
                        name="凡例の要る物")
                label = "灰の点線の箱 = 本当は要るが、この形ではできない物"
            elif k == "graycross":
                d.cross(x + 17, yy + 9, size=16, color=("tx1", 50, 50))
                label = "灰の ✕ = この形では起きない"
            elif k == "red":
                badge(d, x + 10, yy + 9, "N", size=18)
                label = "赤い丸の数字 = できない所（数字は下の比べの表の行）"
            elif k == "gold":
                badge(d, x + 10, yy + 9, "N", gold=True, size=18)
                label = "金の丸の数字 = 確実でない所・作り直しが要る所（数字は表の行）"
            elif k == "callout":
                d.shape(x, yy + 1, 34, 16, "rect", fill=CALLOUT_FILL, line=CALLOUT_LINE, lw=1, name="凡例の吹き出し")
                label = "淡い黄の吹き出し = 問題点 N（右。線の先が、その話の場所）"
            elif k == "purple":
                d.text(x, yy, 36, 18, "紫の字", size=9, font=UD, color=PURPLE, wrap="none", name="凡例の紫")
                label = "紫の字 = GSD の流れを曲げる設定・決まり"
            d.text(x + 40, yy, 300, 18, label, size=10, font=UD, color="000000", wrap="none", name="凡例の文")
            yy += 20
    return yy


# ---- 共通の骨組み ---------------------------------------------------------------
CMD_COLS = [(390, 174, "/gsd-new-project", "案件を始める"), (576, 168, "/gsd-discuss-phase", "決め事を相談"),
            (756, 296, "/gsd-plan-phase", "計画書を書く"), (1064, 475, "/gsd-execute-phase", "作る")]
BOUNDS = (570, 750, 1058)
LANE = 1076                     # 作るの列で、メインから係を起こす幹
STAGE_NOTE = [[("read: ", {"bold": True}), ("モック（デザインの見本の画面）の値を読み取る", {})],
              [("build: ", {"bold": True}), ("画面を作る", {})],
              [("vrt: ", {"bold": True}), ("見た目の回帰テスト（前に撮った画面と比べて、崩れを見つける）", {})],
              [("ほかの工程（例: link＝つなぎ込み）は、図では省いた", {})]]
PLAN_HEAD = "計画書は工程ごとに、画面 1 つにつき 1 枚。agent_hint は、起こすサブエージェントの名前を書く欄"


def skeleton(d, Y0, title, take, why, frame_h):
    d.text(26, Y0, 1500, 44, title, size=28, font=None, color="000000", wrap="none", name="見出し")
    d.shape(34, Y0 + 47, 76, 26, "roundRect", fill=OK_GREEN if take else SUB, line=None, adj=20000,
            text=[[("採る" if take else "採らない", {"bold": True, "color": "FFFFFF"})]], size=13, font=UD,
            ins=(0, 0, 0, 0), name="見出し（採否）")
    d.text(118, Y0 + 48, 1400, 26, why, size=14, font=UD, color="000000", wrap="none", name="見出し（理由）")
    T = Y0 + 82
    d.shape(34, T, 1517, frame_h, "rect", fill=None, line=("tx2",), lw=2.25, name="枠")
    cmds = []
    for x, w, a, b in CMD_COLS:
        cmds.append(d.shape(x, T + 20, w, 61, "homePlate", fill=GREEN, line=GREEN_LINE, lw=0.75, adj=18000,
                            text=[[(a, {"bold": True})], b], size=11, font=YU, color="000000",
                            ins=(2, 0, 9, 0), name=a))
    d.text(78, T + 78, 150, 24, "人（入力するコマンド）", size=12, font=None, color="000000", wrap="none", name="文字")
    d.connector((56, T + 102), (1539, T + 102), color="A6A6A6", lw=1, tail=False, name="区切り線")
    d.text(78, T + 104, 80, 24, "裏側の処理", size=12, font=None, color="000000", wrap="none", name="文字")
    main = d.shape(390, T + 143, 1149, 41, "roundRect", fill=GREEN, line=GREEN_LINE, adj=8000,
                   text="メインエージェント", size=14, font=UD, color=INK, name="メインエージェント")
    MB = T + 184
    with d.group("コマンドごとの境目"):
        for x in BOUNDS:
            line(d, "boundary", (x, T + 6), (x, T + 141), tail=False)
            line(d, "boundary", (x, MB + 2), (x, T + frame_h - 6), tail=False)
    xs = []
    for c, n in zip(cmds, ["(1)", "(3)", "(4)", "(7)"]):
        x = c.site(2)[0]
        line(d, "cmd", (c, 2), (x, T + 143))
        number(d, x + 4, T + 108, n)
        xs.append(x)
    AY = MB + 44
    road = box(d, xs[0] - 52.5, AY, 105, 82, [[("サブエージェント", {"bold": True})], "gsd-roadmapper", "（道のり案内）"])
    line(d, "spawn", (xs[0], MB), (road, 0))
    number(d, xs[0] + 4, MB + 6, "(2)")
    return T, main, MB, AY


def planner(d, MB, AY):
    p = box(d, 764, AY, 120, 82, [[("サブエージェント", {"bold": True})], "gsd-planner", "（計画立案）"])
    line(d, "spawn", (824, MB), (p, 0))
    number(d, 828, MB + 6, "(5)")
    number(d, 846, AY + 86, "(6)")
    return p


def plan_cards(d, ys, lines_of, p):
    d.text(908, ys[0] - 74, 146, 62, PLAN_HEAD, size=9, font=UD, color="000000", name="計画書の説明")
    cards = []
    for y, st in zip(ys, ["read", "build", "vrt"]):
        c = stack(d, 905, y, 140, 58, lines_of(st), name=f"計画書（{st}）")
        line(d, "write", (p, 3), (c, 1))
        cards.append(c)
    return cards


def stage_note(d, y):
    return note(d, 905, y, 150, STAGE_NOTE, size=8.5, color=SUB, name="工程の意味")


def config_read(d, main, cfg):
    line(d, "read", (main, 1), (cfg, 0))


# ---- 図 1 -----------------------------------------------------------------------
def fig1(Y0):
    """絶対の理想の形を、GSD の既定の係だけで組んだ図。書き込む注意は、GSD ではできない所（赤）だけ。
    赤にする所は、GSD の原文（agents/gsd-executor.md・execute-phase の手順書・FEATURES.md）で確かめた物に限る。"""
    d = Drawing(tag="v4fig1", name="図1 理想の形")
    FH = 870
    T, main, MB, AY = skeleton(d, Y0, "図1　理想の形（方法論のやりたい事を、GSD の既定のサブエージェントだけで組む）", False,
                               "GSD の既定の係だけでは、点検・往復・役目ごとの見張り・画面を押す確かめが組めないため（赤い所）", FH)
    d.text(52, T + 6, 334, 70,
           "この図: 方法論が本当にやりたい事を、GSD の既定のサブエージェント（gsd-executor）だけで組んだ形。\n"
           "赤い所が、GSD の仕組みのためにできない所。",
           size=9.5, font=UD, color="000000", name="この図の説明")
    cfg = doc(d, 56, MB + 30, 320, 92,
              [".planning/config.json（GSD の設定）", '"agent_skills": {',
               [('    "gsd-planner":  [方法論の手引き]', {"color": PURPLE})],
               [('    "gsd-executor": [方法論の手引き]', {"color": PURPLE})], "}"],
              size=10, name="config.json", align="l")
    config_read(d, main, cfg)
    p = planner(d, MB, AY)
    man = doc(d, 754, AY + 112, 128, 104,
              [[("方法論の手引き", {"bold": True})], [("（SKILL.md）", {"bold": True})],
               [("agent_skills で渡す", {"color": PURPLE})],
               [("「工程ごとに、画面 1 つにつき 1 枚の計画書を書け」", {"color": PURPLE})]],
              name="計画立案が読む手引き")
    line(d, "read", (p, 2), (man, 0))

    # 作るの列: 理想では、役目の違う 5 体が働く。GSD の既定の係だけなので、5 体とも gsd-executor
    E1 = MB + 96
    GAP = 100
    roles = [["役目: モックの値を読み取る"], ["役目: 1 体目の読み取りを点検する"], ["役目: 画面のコードを書く"],
             ["役目: 撮ってモックと比べ、", "画面を押して確かめる。", "ずれは 3 体目に返して直させる"],
             ["役目: 見た目の回帰テストを書く"]]
    hs = [78 if len(r) > 1 else 54 for r in roles]
    eys = [E1]
    for h in hs[:-1]:
        eys.append(eys[-1] + h + 46)
    FT = E1 - 50
    fr = frame(d, 1086, FT, 222, eys[-1] + hs[-1] + 12 - FT, name="係の枠")
    d.text(1110, FT + 4, 196, 16, "どの役目でも、起きるのは gsd-executor", size=9.5, font=UD, color="000000",
           wrap="none", name="枠の見出し")
    badge(d, 1098, FT + 12, 1)
    exes = []
    for k, role in enumerate(roles):
        lines = [[("GSDのサブエージェント", {"bold": True})], f"gsd-executor（{k + 1} 体目）"] + role
        if k in (1, 3):   # GSD では起こせない係は、灰の点線の箱（本当は要るが、この形ではできない物）
            e = ghost(d, 1098, eys[k], 196, hs[k], lines, name=f"gsd-executor（{k + 1} 体目・起こせない）")
        else:
            e = stack(d, 1098, eys[k], 196, hs[k], lines, name=f"gsd-executor（{k + 1} 体目）")
        exes.append(e)
    ys = [(exes[0].cy + exes[1].cy) / 2 - 29, (exes[2].cy + exes[3].cy) / 2 - 29, exes[4].cy - 29]
    cards = plan_cards(d, ys, lambda st: [[(f"計画書 PLAN.md（{st}）", {"bold": True})], "（agent_hint は無し）"], p)
    stage_note(d, ys[2] + 66)
    for k, e in enumerate(exes):
        line(d, "read", (e, 1), (cards[[0, 0, 1, 1, 2][k]], 3))
    # 起こす: 幹から段ごとに枝。2 体目（点検）と 4 体目（確かめて返す）は、GSD では起こせない（赤の点線と ✕）
    bys = [e.y - 12 for e in exes]
    line(d, "spawn", (LANE, MB), (LANE, bys[-1]), tail=False)
    for k, (e, by) in enumerate(zip(exes, bys)):
        if k in (1, 3):
            color, lw, dash, _, _ = LINES["cannot"]
            d.connector((LANE, by), (e.cx, by), color=color, lw=lw, dash=dash, tail=False, name="できない線")
            line(d, "cannot", (e.cx, by), (e, 0))
            cross(d, LANE + 34, by, 16, RED)
            badge(d, LANE + 58, by, 3 if k == 1 else 4)
        else:
            d.connector((LANE, by), (e.cx, by), color=LINES["spawn"][0], lw=AW, tail=False, name="起こす矢印の枝")
            line(d, "spawn", (e.cx, by), (e, 0))
        number(d, e.cx + 8, by - 24, f"({8 + k})")
    badge(d, exes[1].x + exes[1].w - 4, exes[1].y + 4, 6)
    badge(d, exes[3].x + exes[3].w - 4, exes[3].y + 4, 7)
    # 枠の下: 手引きを読む線と、理想では要る見張り
    FB = fr.y + fr.h
    man2 = doc(d, 1098, FB + 34, 150, 62,
               [[("方法論の手引き（SKILL.md）", {"bold": True})], [("agent_skills で gsd-executor に渡す", {"color": PURPLE})]],
               name="作るときに読む手引き")
    line(d, "read", (1173, FB), (man2, 0))
    g = ghost(d, 1300, FB + 34, 234, 46, ["見張り: 役目ごとに、", "読める物・書ける物を止める"], name="本当は要る見張り")
    line(d, "cannot", (1280, FB), (g, 0))
    cross(d, 1290 + (g.cx - 1290) * 0.45, FB + 15, 20, RED)
    badge(d, 1366, FB + 14, 5)
    legend(d, 56, MB + 140, ["cmd", "spawn", "write", "read", "cannot", "boundary"],
           ["colors", "stack", "doc", "frame", "ghost", "red", "callout", "purple"])

    def rng(e, a, b):
        return (e.y + a, e.y + b)
    items = [
        dict(nums=[1], gold=[False], head="問題点 1　どの役目でも、起きるのは gsd-executor",
             body=[P("GSD は、計画書に agent_hint が無いと、どの計画書でも gsd-executor を起こす。"),
                   P(("役目ごとに道具と手順の違う係を、GSD の既定の係だけでは用意できない。",))],
             to=(fr.x + fr.w, FT + 12), yrange=(FT + 6, exes[0].y - 16), keys=["exec_routing", "planmd"]),
        dict(nums=[3], gold=[False], head="問題点 3　点検する役を、別に起こせない",
             body=[P("GSD の手順書は、計画書 1 本につき gsd-executor を 1 体だけ起こす。"
                     "gsd-executor にはサブエージェントを起こす道具（Agent ツール）が無い。"),
                   P(("作った物を、別の係に点検させられない。",))],
             to=(fr.x + fr.w, exes[1].y + 10), yrange=rng(exes[1], 4, 26), keys=["exec_routing", "tools"]),
        dict(nums=[6], gold=[False], head="問題点 6　gsd-executor の決まりと、点検の役目がぶつかる",
             body=[P("gsd-executor の決まりは「作業ごとにコミットし、SUMMARY.md（作業の記録）を作り、見つけた不具合は自分で直す」。"),
                   P(("点検だけして直さない役にできない。",))],
             to=(fr.x + fr.w, exes[1].y + 40), yrange=rng(exes[1], 30, 52), keys=["executor_md"]),
        dict(nums=[4], gold=[False], head="問題点 4　作る⇄確かめるを、往復させられない",
             body=[P("確かめる役を起こす手順も、ずれを作る役へ返して直させる手順も、GSD の手順書に無い。"),
                   P(("見た目のずれを、別の係が確かめて直させる繰り返しが組めない。",))],
             to=(fr.x + fr.w, exes[3].y + 12), yrange=rng(exes[3], 4, 30), keys=["exec_routing"]),
        dict(nums=[7], gold=[False], head="問題点 7　画面を押して確かめる道具を、持たせられない",
             body=[P("gsd-executor には Playwright MCP（ブラウザを動かす道具）が無い。"
                     "GSD に足す部品（Capability）でも、起こすときの指定でも足せない。"),
                   P(("画面を押して確かめる役にできない。",))],
             to=(fr.x + fr.w, exes[3].y + 52), yrange=rng(exes[3], 40, 74), keys=["tools"]),
        dict(nums=[5], gold=[False], head="問題点 5　役目ごとの見張りを掛けられない",
             body=[P("見張り（Claude Code の hook）は、サブエージェントの名前で相手を決める。GSD が起こす名前は、どの役目でも gsd-executor。"),
                   P(("役目ごとに、読める物・書ける物を止められない。",))],
             to=(g.x + g.w, g.cy), yrange=(g.y + 6, g.y + g.h - 6), keys=["exec_routing"]),
    ]
    bottom = place_problems(d, items, MB - 16)
    return d, max(T + FH, bottom)


# ---- 図 2 -----------------------------------------------------------------------
def fig2(Y0):
    d = Drawing(tag="v4fig2", name="図2 Capability に全振り")
    FH = 860
    T, main, MB, AY = skeleton(d, Y0, "図2　Capability に全振りした形（理想になるべく近づける）", False,
                               "起こすサブエージェントを選べず、点検も往復も別のサブエージェントに任せられないため", FH)
    d.text(52, T + 6, 334, 70,
           "考えた理由: 方法論独自のサブエージェントを案件に置くと、GSD の版上げのたびにキット（方法論の部品一式）を直す手間がかかる。\n"
           "Capability（GSD に部品を足す正式な仕組み）に全部載せれば、この手間が減る。",
           size=9.5, font=UD, color="000000", name="考えた理由")
    cap = doc(d, 56, MB + 30, 320, 78,
              [".gsd/capabilities/sdd-kit/capability.json",
               [('"contributions": 依頼文に書き足す文', {"color": PURPLE})],
               [('"skills": 方法論の手引き SKILL.md', {"color": PURPLE})],
               '"gates" / "hooks": 自動の検査と見張り（図3 でも使う）'],
              size=10, name="Capability", align="l")
    config_read(d, main, cap)
    p = planner(d, MB, AY)
    add = doc(d, 754, AY + 112, 128, 104,
              [[("gsd-planner への", {"bold": True})], [("依頼文に書き足す文", {"bold": True})],
               [("「計画書に工程（sdd.stage）を書け」", {"color": PURPLE})]], name="計画立案への書き足し")
    line(d, "read", (p, 2), (add, 0))

    prompt = doc(d, 1088, MB + 32, 300, 96,
                 [[("gsd-executor への依頼文に書き足す文", {"bold": True})],
                  [("「sdd.stage を読み、その工程の手引き（Capability の中の", {"color": PURPLE})],
                  [("SKILL.md）を読んでから作業せよ」", {"color": PURPLE})],
                  [("（メインが 3 体すべての依頼文に書き足す）", {"color": SUB})]],
                 size=9.5, name="gsd-executor への書き足し")
    line(d, "spawn", (prompt.cx, MB), (prompt, 0))
    number(d, prompt.cx + 6, MB + 2, "(8)")
    badge(d, prompt.x + prompt.w + 14, prompt.y + 16, 2, gold=True)
    badge(d, prompt.x + prompt.w + 14, prompt.y + 42, 8, gold=True)
    FT = prompt.y + prompt.h + 56
    E1 = FT + 50
    GAP = 124
    fr = frame(d, 1086, FT, 186, 50 + 2 * GAP + 58 + 40, name="係の枠")
    d.text(1110, FT + 4, 160, 30, "どの計画書でも、起きるのは\ngsd-executor（選べない）", size=9.5, font=UD,
           color="000000", name="枠の見出し")
    badge(d, 1098, FT + 12, 1)
    exes = []
    for k in range(3):
        exes.append(stack(d, 1098, E1 + GAP * k, 150, 58,
                          [[("GSDのサブエージェント", {"bold": True})], f"gsd-executor（{k + 1} 体目）",
                           [("＋書き足された文", {"color": PURPLE})]], name=f"gsd-executor（{k + 1} 体目）"))
    cards = plan_cards(d, [e.y for e in exes],
                       lambda st: [[(f"計画書 PLAN.md（{st}）", {"bold": True})], [(f"sdd.stage: {st}", {"color": PURPLE})],
                                   "（agent_hint は無し）"], p)
    stage_note(d, exes[2].y + 66)
    for e, c in zip(exes, cards):
        line(d, "read", (e, 1), (c, 3))
    bys = [e.y - 10 for e in exes]
    line(d, "spawn", (prompt, 1), (LANE, prompt.cy), tail=False)
    d.connector((LANE, prompt.cy), (LANE, bys[-1]), color=LINES["spawn"][0], lw=AW, tail=False, name="起こす矢印の幹")
    for e, by in zip(exes, bys):
        d.connector((LANE, by), (e.cx, by), color=LINES["spawn"][0], lw=AW, tail=False, name="起こす矢印の枝")
        line(d, "spawn", (e.cx, by), (e, 0))
    wants = [("今の方法論: 読み取る役と、", "点検する役を別々に起こす", 3),
             ("今の方法論: 作る役と、撮って比べる役を", "往復させて直す", 4)]
    ghosts = []
    for k, (a, b, n) in enumerate(wants):
        g = ghost(d, 1304, exes[k].y + 6, 230, 46, [a, b], name=f"本当は要る物（{n}）")
        line(d, "cannot", (exes[k], 3), (g, 1))
        mx = (exes[k].x + exes[k].w + g.x) / 2 + 2
        cross(d, mx, g.cy, 18, RED)
        badge(d, mx, g.cy - 20, n)
        ghosts.append(g)
    FB = fr.y + fr.h
    d.text(1110, FB - 32, 160, 28, "どの gsd-executor にも\nPlaywright MCP が無い", size=9, font=UD, color="000000",
           name="枠の下の注")
    badge(d, 1098, FB - 22, 7, gold=True)
    w = ghost(d, 1086, FB + 34, 220, 46, ["見張り: 役目ごとに、", "読める物・書ける物を止める"], name="本当は要る見張り")
    line(d, "cannot", (fr, 2), (w, 0))
    cross(d, fr.cx, FB + 17, 18, RED)
    badge(d, fr.cx + 20, FB + 17, 5)
    yy = legend(d, 56, MB + 126, ["cmd", "spawn", "write", "read", "cannot", "boundary"],
                ["colors", "stack", "doc", "frame", "ghost", "red", "gold", "callout", "purple"])
    with d.group("位置づけ"):
        e = note(d, 56, yy + 14, 330,
                 "GSD が正式に用意した、計画書ごとに起こすサブエージェントを選ぶ仕組み（agent_hint）を使わず、依頼文の書き足しに頼る形。",
                 size=10, name="位置づけの文")
        link_tags(d, 56, e + 2, ["1689", "4740", "1143"])

    items = [
        dict(nums=[2], gold=[True], head="問題点 2　依頼文への書き足しが、入る保証が無い",
             body=[P("書き足しは、メインが手順書を読んで自分で入れる仕組み。"),
                   P(("入らないと、gsd-executor は方法論の手順を知らずに作る。",))],
             to=(prompt.x + prompt.w + 24, prompt.y + 16), keys=["4139", "step275"]),
        dict(nums=[8], gold=[True], head="問題点 8　手引きの置き場所を、依頼文で教えるしかない",
             body=[P("Capability の手引き（SKILL.md）は、サブエージェントが手引きを探す .claude/skills/ に写されない（GSD 1.20.0 で確認）。"),
                   P(("教え漏れると、手引きを読まずに作る。",))],
             to=(prompt.x + prompt.w + 24, prompt.y + 42), keys=["ship"]),
        dict(nums=[1], gold=[False], head="問題点 1　どの計画書でも gsd-executor が起きる",
             body=[P("Capability には、起こすサブエージェントを指定する手段が無い。"),
                   P(("工程に要る道具と手順を持つサブエージェントに、任せられない。",)),
                   P("別の案（Capability の feature で、gsd-executor の前後に処理を足す）でも、gsd-executor は必ず動き、"
                     "欲しくないコードまで作る。"),
                   P("Capability に書いたサブエージェントは .claude/agents/ に置かれず、起こせない（直す修正は承認済み・まだ出ていない）。")],
             to=(fr.x + fr.w, FT + 14), yrange=[(FT + 8, exes[0].y), (exes[0].y + 60, exes[1].y)],
             keys=["1689", "planmd", "2363", "manifest", "5738"]),
        dict(nums=[3], gold=[False], head="問題点 3　点検を、別のサブエージェントに任せられない",
             body=[P("gsd-executor はサブエージェントを起こせない（Agent ツール〔サブエージェントを起こす道具〕が無く、"
                     "Capability でも足せない）。"),
                   P(("作った本人が点検するしかなく、見落としが残る。",))],
             to=(ghosts[0].x + ghosts[0].w, ghosts[0].cy), yrange=(ghosts[0].y + 6, ghosts[0].y + ghosts[0].h - 6), keys=["tools", "firstparty"]),
        dict(nums=[4], gold=[False], head="問題点 4　作る→比べる→直すを、往復させる相手がいない",
             body=[P("1 体の gsd-executor が、作って、比べて、直すしかない。"),
                   P(("見た目のずれを、別の目で確かめずに合格にしてしまう。",))],
             to=(ghosts[1].x + ghosts[1].w, ghosts[1].cy), yrange=(ghosts[1].y + 6, ghosts[1].y + ghosts[1].h - 6), keys=["tools"]),
        dict(nums=[7], gold=[True], head="問題点 7　画面を押して確かめる道具が無い",
             body=[P("図1 と同じく、gsd-executor には Playwright MCP（ブラウザを動かす道具）が無く、Capability でも足せない。"),
                   P(("押して画面の移り方を確かめる手順は、作り直しが要る。",))],
             to=(fr.x + fr.w, FB - 20), yrange=(exes[2].y + 66, FB - 6), keys=["tools"]),
        dict(nums=[5], gold=[False], head="問題点 5　テストを書く役と実装を書く役を、見張りで分けられない",
             body=[P("1 体の gsd-executor がテストも実装も書くので、名前で分けようが無い。"),
                   P(("テストを実装に合わせて弱めても、見つからない。",))],
             to=(w.x + w.w, w.cy), yrange=(w.y + 6, w.y + w.h - 6), keys=[]),
    ]
    bottom = place_problems(d, items, MB - 16)
    return d, max(T + FH, bottom)


# ---- 図 3 -----------------------------------------------------------------------
def fig3(Y0):
    d = Drawing(tag="v4fig3", name="図3 今の方法論の形")
    FH = 830
    T, main, MB, AY = skeleton(d, Y0, "図3　今の方法論の形（agent_hint で、方法論独自サブエージェントを起こす）", True,
                               "工程ごとに合ったサブエージェントを起こせて、役目ごとの読み書きも見張りで止められるため", FH)
    cfg = doc(d, 56, MB + 30, 320, 108,
              [".planning/config.json（GSD の設定）", '"agent_skills": {',
               [('    "gsd-planner":  [".claude/skills/sdd-planner"]', {"color": PURPLE})], "}",
               [('"agent_hint_routing": true', {"color": PURPLE})],
               [("（計画書の agent_hint で、起こすサブエージェントを選ぶ設定）", {"color": SUB})]],
              size=10, name="config.json", align="l")
    config_read(d, main, cfg)
    with d.group("位置づけ"):
        e = note(d, 56, cfg.y + cfg.h + 12, 330,
                 "計画書の agent_hint で起こすサブエージェントを選ぶ仕組みと、\n"
                 "agent_skills で手引きを渡す仕組みは、GSD が正式に用意している。", size=10, name="位置づけの文")
        link_tags(d, 56, e + 2, ["planmd_603", "skills_603", "routing", "settings"])
    p = planner(d, MB, AY)
    man = doc(d, 754, AY + 112, 128, 124,
              [[("sdd-planner", {"bold": True})], [("（SKILL.md）", {"bold": True})],
               [("「PLAN.md の agent_hint に、その工程の方法論独自サブエージェントの名前を書け」", {"color": PURPLE})]],
              name="計画立案が読む手引き")
    line(d, "read", (p, 2), (man, 0))

    note(d, 1112, MB + 30, 424,
         [[("方法論独自サブエージェント sdd-relay-○○", {"bold": True}), ("（.claude/agents/ に置く）: 工程ごとに 1 種類あり、", {})],
          [("子のサブエージェントを順に起こして作業させる", {})]], size=9.5, name="方法論独自サブエージェントの説明")
    number(d, LANE + 4, MB + 4, "(8)")
    FT = MB + 76
    R0 = FT + 18
    R1 = R0 + 58 + 28
    R2 = R1 + 104 + 28
    rows = [(R0, 58), (R1, 104), (R2, 58)]
    fr = frame(d, 1086, FT, 452, R2 + 58 + 14 - FT, name="係の枠")
    relays = []
    for (ry, rh), st in zip(rows, ["read", "build", "vrt"]):
        relays.append(stack(d, 1098, ry + (rh - 58) / 2, 170, 58,
                            [[("方法論独自サブエージェント", {"bold": True})], [(f"sdd-relay-{st}", {"color": PURPLE})],
                             "（計画書 1 枚に 1 体）"], fill=BLUE, line=BLUE_LINE, name=f"方法論独自（{st}）"))
    c_read = box(d, 1310, R0, 222, 58,
                 [[("子のサブエージェント", {"bold": True})], "dna-extractor",
                  [("（モックから色・字・余白の値を読み取る。別の回で点検）", {"size": 8.5})]],
                 fill=BLUE, line=BLUE_LINE, name="子（read）")
    with d.group("子（build）"):
        cf = frame(d, 1310, R1, 222, 104, name="子の枠")
        imp = box(d, 1316, R1 + 6, 94, 40, [[("implementer", {"bold": True})], "（作る）"],
                  fill=BLUE, line=BLUE_LINE, name="implementer")
        d.text(1410, R1 + 14, 22, 22, "⇄", size=14, font=UD, color="000000", align="ctr", name="往復の印")
        ver = box(d, 1432, R1 + 6, 94, 40, [[("visual-verifier", {"bold": True})], "（確かめる）"],
                  fill=BLUE, line=BLUE_LINE, name="visual-verifier")
        note(d, 1318, R1 + 52, 210,
             "確かめ方: build では撮影スクリプト（shoot.mjs）で撮って比べる。\n"
             "工程 link（つなぎ込み）では、Playwright MCP で画面を押して確かめる", size=8.5, name="確かめ方の注")
    c_vrt = box(d, 1310, R2, 222, 58,
                [[("子のサブエージェント", {"bold": True})], "vrt-test-implementer",
                 [("（見た目の回帰テストを書く）", {"size": 8.5})]],
                fill=BLUE, line=BLUE_LINE, name="子（vrt）")
    for r, c in zip(relays, [c_read, cf, c_vrt]):
        line(d, "spawn", (r, 3), (c, 1))
    number(d, 1280, relays[0].cy - 28, "(9)")
    cards = plan_cards(d, [r.y for r in relays],
                       lambda st: [[(f"計画書 PLAN.md（{st}）", {"bold": True})], [("agent_hint:", {"color": PURPLE})],
                                   [(f"sdd-relay-{st}", {"color": PURPLE})]], p)
    for r, c in zip(relays, cards):
        line(d, "read", (r, 1), (c, 3))
    nb = stage_note(d, cards[2].y + 66)
    pf = frame(d, 899, cards[0].y - 8, 157, nb + 6 - cards[0].y + 8, name="計画書の枠")
    FB = fr.y + fr.h
    ex = box(d, 1120, FB + 22, 150, 52, [[("GSDのサブエージェント", {"bold": True})], "gsd-executor"],
             name="起きない gsd-executor")
    bys = [r.y - 9 for r in relays]
    line(d, "spawn", (LANE, MB), (LANE, ex.cy), tail=False)
    for r, by in zip(relays, bys):
        d.connector((LANE, by), (r.cx, by), color=LINES["spawn"][0], lw=AW, tail=False, name="起こす矢印の枝")
        line(d, "spawn", (r.cx, by), (r, 0))
    line(d, "spawn", (LANE, ex.cy), (ex, 1))
    d.cross(1098, ex.cy, size=18, color=("tx1", 50, 50))
    note(d, 1280, FB + 18, 176,
         "計画書に agent_hint があるので、GSD 既定の gsd-executor は起きず、agent_hint のサブエージェントが起きる",
         size=9, name="起きない訳")
    wy = FB + 98
    watch = auto_check(d, 1086, wy, 452, 62,
                       [[("自動の検査: キットの見張り（Claude Code の hook）", {"bold": True})],
                        "サブエージェントの名前ごとに、読める物・書ける物・起こせるサブエージェントを止める。",
                        "例: 確かめ係（visual-verifier）と規約の点検係は、実装のコードとテストを書けない"],
                       name="キットの見張り")
    line(d, "read", (1490, wy), (1490, FB))
    gate = auto_check(d, 760, wy + 86, 290, 60,
                      [[("自動の検査: Capability の門", {"bold": True})],
                       "計画の後と、まとまりごとに検査する。",
                       "次のまとまりを続けて流させる書き足しも 1 つ入れる"], name="Capability の門")
    line(d, "read", (pf.x + pf.w / 2, gate.y), (pf, 2))
    badge(d, gate.x + gate.w - 6, gate.y, 2, gold=True)
    badge(d, fr.x + fr.w - 4, fr.y, 9, gold=True)
    legend(d, 56, e + 40, ["cmd", "spawn", "write", "read", "boundary"],
           ["colors", "stack", "doc", "frame", "auto", "graycross", "gold", "callout", "purple"])

    items = [
        dict(nums=[9], gold=[True], head="問題点 9　引き受けた手間: 方法論独自サブエージェントが増える",
             body=[P("sdd-relay-○○ は工程ごとに 1 種類で、全部で 14 種類（図は read・build・vrt の 3 つだけ描いた）。"),
                   P("BFF〔画面のための裏側の窓口〕の部品を外すと 11 種類。"),
                   P("増えた分だけ、宣言の表（係ごとに読める物・書ける物を書いた表）と実体のずれを見張る手間と、",
                     ("GSD の版上げのたびの確かめ直しが要る。",))],
             to=(fr.x + fr.w, fr.y + 30), yrange=(fr.y + 10, fr.y + fr.h - 10), keys=["2363"]),
        dict(nums=[2], gold=[True], head="問題点 2　次のまとまりを続けて流させる書き足しは、入る保証が無い",
             body=[P("GSD は計画書を何本かずつの一区切り（まとまり）で流す。この書き足しは、メインが依頼文に入れる仕組み。"),
                   P(("入らないと途中で止まる",), "（同じコマンドを打ち直せば、続きから流れる）。")],
             to=(gate.x + gate.w, gate.cy), yrange=(gate.y + 8, gate.y + gate.h - 8), keys=["4139"]),
    ]
    bottom = place_problems(d, items, MB - 16)
    return d, max(T + FH, bottom)


# ---- 比べの表 ---------------------------------------------------------------------
X, G, O = RED, GOLD, OK_GREEN
ROWS = [
    ("工程ごとに、合った道具と手順を持つサブエージェントを起こす",
     (X, "✕ 起きるのはいつも gsd-executor。役目ごとに道具の違う係を起こせない"),
     (X, "✕ Capability には、起こすサブエージェントを指定する手段が無い"),
     (O, "○ 計画書の agent_hint に書いたサブエージェントを、GSD が起こす"),
     "工程に要る道具・手順を持つサブエージェントに任せられない", ["1689", "planmd"]),
    ("方法論の手順を、作業するサブエージェントに確実に伝える",
     (O, "○ agent_skills の手引きで、方法論の手順を gsd-executor に渡せる"),
     (G, "△ メインの書き足しを当てにするしかなく、入る保証が無い"),
     (G, "△ 手順はサブエージェントのファイルにあり、起きた時点で持つ。ただし次のまとまりを流させる書き足し 1 つは保証が無い"),
     "入らないと、方法論の手順を知らずに作る（図3 では途中で止まる）", ["4139", "step275"]),
    ("作った物を、別のサブエージェントに点検させる",
     (X, "✕ 計画書 1 本に 1 体だけ。gsd-executor は点検役を起こせない"),
     (X, "✕ gsd-executor はサブエージェントを起こせない"),
     (O, "○ 方法論独自サブエージェントが子を起こし、点検させる"),
     "作った本人が点検するしかなく、見落としが残る", ["tools", "firstparty"]),
    ("作る→点検→直すの繰り返しを、GSD の手順書どおりに回す",
     (X, "✕ 確かめる役を起こし、ずれを返して直させる手順が GSD に無い"),
     (X, "✕ 往復させる相手がいない"),
     (O, "○ 方法論独自サブエージェントが回す。手順はそのファイルにある"),
     "手順書に無い動きを、メインが毎回守る保証が無い", ["exec_routing"]),
    ("役目ごとに、読める物・書ける物を見張りで止める",
     (X, "✕ 見張りから見える名前は、いつも gsd-executor"),
     (X, "✕ gsd-executor 1 体が全部するので、分けようが無い"),
     (O, "○ 画像・テスト・実装のコード・書類を、サブエージェントの名前で見張りが止める"),
     "依頼文の決まりだけだと、破られても見つからない", []),
    ("GSD のサブエージェントの決まりと、ぶつからない",
     (X, "✕ 「作業ごとにコミット・SUMMARY.md を作る・不具合は自分で直す」の決まりで、点検だけの役にできない"),
     (O, "○ 計画書 1 本を 1 体に最後まで作らせるので、決まりどおり"),
     (O, "○ gsd-executor は起こさない"),
     "どちらが勝つか分からず、点検役がコードを直してしまうことがある", ["executor_md"]),
    ("画面を押して確かめる（つなぎ込みの工程 link）",
     (X, "✕ Playwright MCP を持たせられない（Capability でも足せない）"),
     (G, "△ Playwright MCP が無い。押して確かめる手順は作り直し"),
     (O, ["○ 撮って比べるのは撮影スクリプト（shoot.mjs）。",
          "押して確かめるのは、Playwright MCP を持つ子（visual-verifier）"]),
     "押して確かめる手順を、作り直す必要がある", ["tools"]),
    ("手引き（SKILL.md）を、サブエージェントが読む場所に置く",
     (O, "○ config.json の agent_skills で、gsd-executor に渡せる"),
     (G, "△ Capability の手引きは .claude/skills/ に写されない（GSD 1.20.0 で確認）"),
     (O, "○ 導入の手順（gsd-setup）が .claude/skills/ に置く"),
     "置き場所を依頼文で教えるしかない", ["ship", "skills_603"]),
    ("GSD の版上げで、直す手間が少ない",
     (O, "○ GSD の既定のサブエージェントだけなので、直す物が少ない"),
     (O, "○ 直すのは Capability の 1 か所だけ"),
     (G, "△ 版上げのたびに、方法論独自サブエージェント・GSD の手順書への直し・見張りが、新しい版でも合うかを確かめ直す"),
     "版上げのたびに手間が掛かる", ["2363"]),
]
COLS = [("要ること", 270), ("図1 理想の形", 270), ("図2 Capability に全振り", 270), ("図3 今の方法論の形", 290),
        ("できないと何が困るか", 250), ("根拠", 100)]


def table(Y0):
    d = Drawing(tag="v4table", name="比べの表")
    d.text(26, Y0, 1500, 40, "比べの表（要ることごとに、3 つの形を並べる。行の番号は、図の丸の数字と同じ）",
           size=24, font=None, color="000000", wrap="none", name="見出し")
    x0, y = 34, Y0 + 48
    with d.group("一覧表"):
        cx = x0
        for title, w in COLS:
            d.shape(cx, y, w, 30, "rect", fill=("bg2",), line="A6A6A6", lw=0.75,
                    text=[[(title, {"bold": True})]], size=11, font=UD, color="000000", name="表の見出し")
            cx += w
        y += 30
        for r, (need, f1, f2, f3, trouble, keys) in enumerate(ROWS):
            cells = [(INK, need), f1, f2, f3, (INK, trouble)]
            hs = [need_h(t, w, 10, UD, (30 if c == 0 else 6, 2, 4, 2)) for c, ((_, t), (_, w)) in enumerate(zip(cells, COLS))]
            rh = max(40, max(hs) + 4, 20 * len(keys) + 6)
            cx = x0
            for c, ((color, t), (_, w)) in enumerate(zip(cells, COLS)):
                d.shape(cx, y, w, rh, "rect", fill="FFFFFF", line="A6A6A6", lw=0.75, text=t, size=10, font=UD,
                        color=color, align="l", anchor="ctr", ins=(30 if c == 0 else 6, 2, 4, 2), name="表のセル")
                if c == 0:
                    badge(d, cx + 14, y + rh / 2, r + 1)
                cx += w
            d.shape(cx, y, COLS[-1][1], rh, "rect", fill="FFFFFF", line="A6A6A6", lw=0.75, name="表のセル")
            for k, key in enumerate(keys):
                d.text(cx + 4, y + 3 + k * 20, 78, 20, f"参考リンク{k + 1}", size=11, font=UD, color=LINK,
                       underline=True, link=L[key], wrap="none", name=f"参考リンク{k + 1}")
            y += rh
    with d.group("結論"):
        d.shape(34, y + 14, 1450, 48, "rect", fill=("accent2", 20, 80), line=("accent2",), lw=1,
                text=[[("結論: ", {"bold": True}),
                       ("図1 の理想のうち ✕ の所（1・3〜7）は、GSD の仕組みで決まっていて、方法論の側では変えられない。"
                        "図2 でも多くは ✕ のまま。", {})],
                      [("→ 図3 の形（agent_hint で方法論独自サブエージェントを起こす）で進める。"
                        "図3 でも 2 は △ が残り、9 は引き受けた手間。", {})]],
                size=11, font=UD, color="000000", align="l", anchor="ctr", ins=(10, 2, 8, 2), name="結論")
        d.text(34, y + 74, 1300, 34,
               "先の見込み: 「Capability に入れたサブエージェントも GSD が案件に置く」修正は、承認済みで、まだ出ていない。\n"
               "出れば、サブエージェントの配り方は Capability に寄せられる（選ぶのは agent_hint のままで、表の 1 は変わらない）。",
               size=10, font=UD, color="000000", name="先の見込み")
        d.text(1360, y + 80, 78, 20, "参考リンク1", size=11, font=UD, color=LINK, underline=True, link=L["5738"],
               wrap="none", name="参考リンク1")
    return d, y + 110


def head():
    d = Drawing(tag="v4head", name="頭書き")
    d.text(26, 0, 1500, 44, "GSD に独自処理を挟む方法 ― 理想の形から、今の方法論の形まで", size=26, font=None,
           color="000000", wrap="none", name="見出し")
    d.text(26, 46, 1500, 24,
           "結論: 図1 の理想の形のうち、赤い所は GSD の仕組みのためにできない。図2 でも埋まらないので、図3 を採る。",
           size=13, font=UD, color="000000", wrap="none", name="見出し（結論）")
    d.text(26, 70, 1500, 22,
           "GSD: AI に計画書を書かせてから作らせる道具。　メイン: 人と話す AI。　"
           "サブエージェント: メインが起こして作業を任せる別の AI。",
           size=11, font=UD, color="000000", wrap="none", name="見出し（用語）")
    return d


# ---- 書き込み ---------------------------------------------------------------------
def strip_old(src, tmp):
    """v3 の図（目印 v2*）とそのリンクを外した写しを作る。人の図（目印の無い図）は v3 に無い。"""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "xl/drawings/drawing1.xml":
                x = data.decode("utf-8")
                x = re.sub(r'<xdr:twoCellAnchor\b(?:(?!</xdr:twoCellAnchor>).)*?descr="xdraw:v2\w*".*?</xdr:twoCellAnchor>',
                           "", x, flags=re.S)
                data = x.encode("utf-8")
            elif info.filename == "xl/drawings/_rels/drawing1.xml.rels":
                data = re.sub(r'<Relationship Id="rIdXv2\w*"[^>]*/>', "", data.decode("utf-8")).encode("utf-8")
            zout.writestr(info, data)


def move_end(grid, end_y):
    """「以上」のセルを、最後の図の下の行へ動かす。"""
    row = int(end_y * 12700 // grid.row_h) + 2

    def edit(sheet):
        sheet = re.sub(r'<row r="\d+" spans="2:2"><c r="B\d+" t="s"><v>0</v></c></row>',
                       f'<row r="{row}" spans="2:2"><c r="B{row}" t="s"><v>0</v></c></row>', sheet)
        return re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="B{row}"/>', sheet)
    return edit


def build_all(src, dst):
    grid = measure_grid(src)
    tmp = dst + ".tmp.xlsx"
    strip_old(src, tmp)
    h = head()
    d1, b1 = fig1(104)
    d2, b2 = fig2(b1 + 50)
    d3, b3 = fig3(b2 + 50)
    dt, bt = table(b3 + 50)
    write_into(tmp, dst, [h, d1, d2, d3, dt], grid=grid, sheet_edit=move_end(grid, bt))
    os.remove(tmp)
    return [(104, b1), (b1 + 50, b2), (b2 + 50, b3), (b3 + 50, bt)]


if __name__ == "__main__":
    spans = build_all(sys.argv[1], sys.argv[2])
    print(f"書きました: {sys.argv[2]}")
    for name, (a, b) in zip(["図1", "図2", "図3", "比べの表"], spans):
        print(f"  {name}: y {a:.0f}〜{b:.0f}")
