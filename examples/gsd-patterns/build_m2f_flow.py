"""運用 m2f（モックから画面を作る）の動きを、「GSDに独自処理を挟む3つの方法.xlsx」と同じ描き方で 1 枚の図にする台本。

使い方: python3 build_m2f.py 見本.xlsx 出力.xlsx
- 見本の xlsx から図とセルの文字を外した空の土台を作り、そこへ描く（セルの大きさは見本の図から測る）。
- 線の種類は凡例の 5 つだけ。凡例も同じ表（LEGEND）から描く。
"""
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.expanduser("~/.claude/skills/excel-figure/scripts"))
from xdraw import Drawing, measure_grid, write_into  # noqa: E402

GREEN = ("accent6", 20, 80)      # GSD の物
GREEN_LINE = ("accent3",)
BLUE = ("accent1", 20, 80)       # 方法論（キット）の物
BLUE_LINE = "2F6DB5"
ARROW = ("accent2",)             # 起こす（橙）
AW = 1.5
GRAY = ("tx1", 50, 50)           # 読む（灰の点線）
CMD = "404040"                   # コマンドで動く（黒）
WRITE = ("accent3",)             # 書く（緑）
NEXT = "7F7F7F"                  # 次の区切りへ（灰の実線）
PURPLE = ("accent5",)
UD = "BIZ UDGothic"
YU = "Yu Gothic"
Y0 = 0
FRAME_W = 2382
FRAME_H = 740


def box(d, x, y, w, h, lines, fill=GREEN, line=GREEN_LINE, size=10, name="箱"):
    return d.shape(x, y, w, h, "roundRect", fill=fill, line=line, lw=0.75, adj=8000, align="ctr",
                   text=lines, size=size, font=YU, ins=(4, 2, 4, 2), name=name)


def stack(d, x, y, w, h, lines, fill=GREEN, line=GREEN_LINE, name="重なった箱"):
    """同じ物が複数あることを示す、重なったカード。前の 1 枚を返す。"""
    with d.group(name):
        for k in (2, 1):
            d.shape(x + 4 * k, y + 3 * k, w, h, "roundRect", fill=fill, line=line, lw=0.75, adj=8000, name="重なり")
        return box(d, x, y, w, h, lines, fill, line, 10, name)


def command(d, a, b):
    d.connector(a, b, color=CMD, lw=AW, name="コマンド")


def spawn(d, a, b, name="起こす矢印"):
    d.connector(a, b, color=ARROW, lw=AW, name=name)


def write(d, a, b):
    d.connector(a, b, color=WRITE, lw=AW, name="書く")


def read(d, a, b):
    d.connector(a, b, color=GRAY, lw=1.25, dash="dash", name="読む")


def number(d, cx, y, label):
    d.text(cx - 22, y, 44, 30, label, size=16, font=None, color="000000", align="ctr", name="番号")


LEGEND = [(CMD, AW, None, "人のコマンドで、メインエージェントが動く"),
          (ARROW, AW, None, "サブエージェントを起こす（呼び出す）"),
          (WRITE, AW, None, "書く（矢印の先の物を作る）"),
          (GRAY, 1.25, "dash", "読んで従う（矢印の先の物を読む）"),
          (NEXT, AW, None, "次の区切りへ（(2) から繰り返す）"),
          ("1F3864", 1.25, "sysDot", "図1 と図2 の対応（点線の枠の中を図2 で詳しく）"),
          (("tx1", 35, 65), 1, "sysDash", "コマンドごとの境目（その上のコマンドで動く範囲）")]


GATE_TAG = "2F5597"


def gate(d, x, y, w, h, lines, name="自動の検査"):
    """自動で動く検査（GSD の差し込み口から呼ばれるキットの道具）。両端に縦の帯がある箱と「自動の検査」の札で、係と見分ける。"""
    with d.group(name):
        d.shape(x, y - 15, 74, 14, "roundRect", fill=GATE_TAG, line=None, adj=30000,
                text=[[("⚙ 自動の検査", {"bold": True})]], size=8, color="FFFFFF", font=YU, ins=(2, 0, 2, 0), name="札（自動の検査）")
        return d.shape(x, y, w, h, "flowChartPredefinedProcess", fill=BLUE, line=GATE_TAG, lw=1.25,
                       text=lines, size=10, font=YU, ins=(2, 2, 2, 2), name=name)


def rule_doc(d, x, y, w, h, lines, fill, name):
    """読むだけの決まり（手順書・CLAUDE.md・手引き）。紙の書類の形と点線の縁で、その場で作る物と分ける。"""
    return d.shape(x, y, w, h, "flowChartDocument", fill=fill, line=("tx1", 35, 65), lw=1, dash="dash",
                   text=lines, size=9, font=YU, ins=(3, 2, 3, 6), name=name)


def legend(d, x, y):
    with d.group("凡例"):
        d.text(x, y - 24, 200, 20, [[("凡例", {"bold": True})]], size=11, font=UD, color="000000", wrap="none")
        for k, (color, lw, dash, label) in enumerate(LEGEND):
            yy = y + 20 * k
            d.connector((x, yy + 9), (x + 34, yy + 9), color=color, lw=lw, dash=dash, name="凡例の線")
            d.text(x + 40, yy, 220, 18, label, size=10, font=UD, color="000000", wrap="none")
        yy = y + 20 * len(LEGEND) + 6
        box(d, x, yy, 34, 18, "", name="凡例の色（GSD）")
        d.text(x + 40, yy, 220, 18, "GSD が持つ物（係・書類）", size=10, font=UD, color="000000", wrap="none")
        box(d, x, yy + 24, 34, 18, "", fill=BLUE, line=BLUE_LINE, name="凡例の色（方法論）")
        d.text(x + 40, yy + 24, 220, 18, "方法論（キット）が持つ物", size=10, font=UD, color="000000", wrap="none")
        d.text(x, yy + 72, 280, 18, [[("紫の字", {"bold": True, "color": "7030A0"}), ("＝キットが GSD の流れを曲げる仕掛け", {})]],
               size=10, font=UD, color="000000", wrap="none", name="凡例（仕掛け）")
        d.shape(x, yy + 48, 34, 18, "flowChartPredefinedProcess", fill=BLUE, line=GATE_TAG, lw=1.25, name="凡例の形（自動の検査）")
        rule_doc(d, x, yy + 96, 34, 20, "", "FFFFFF", name="凡例の形（読むだけの決まり）")
        d.text(x + 40, yy + 97, 260, 18, "読むだけの決まり（手順書・CLAUDE.md・手引き）", size=10, font=UD, color="000000",
               wrap="none")
        d.text(x + 40, yy + 48, 240, 18, "自動で動く検査（GSD の差し込み口から呼ばれる）", size=10, font=UD, color="000000",
               wrap="none")


KID_GROUPS = {
    "read": ("画面ごとに、読む → 監査（2〜3 画面ずつ同時に）",
             [("① dna-extractor", "読む（主担当）"), ("② dna-extractor", "監査（別の回で確かめる）")], "→"),
    "build": ("実装と確かめを別の係に分ける（最大 5 周。共通の build は ① だけ）",
              [("① implementer", "画面のコードを書く"), ("② visual-verifier", "撮ってお手本と比べる")], "⇄"),
    "link": ("表を書く → 画面ごとにつなぐ → 押して確かめる",
             [("① dna-extractor", "つなぎ表"), ("② implementer", "画面ごとにつなぐ"), ("③ visual-verifier", "押して確かめる")], "→"),
}


def kid_group(d, x, y, w, h, st):
    """方法論の係が順に起こす子の係のまとまり（点線の枠）。中に、起こす順の番号つきで子の係を並べる。"""
    label, kids, sep = KID_GROUPS[st]
    with d.group(f"子の係（{st}）"):
        if st == "build":   # 方法論の係（画面ごと）ごとに 1 つずつある
            for k in (2, 1):
                d.shape(x + 4 * k, y + 3 * k, w, h, "rect", fill="F3F7FC", line=BLUE_LINE, lw=1, dash="sysDash", name="重なり")
        frame = d.shape(x, y, w, h, "rect", fill="F3F7FC", line=BLUE_LINE, lw=1, dash="sysDash", name=f"子の係の枠（{st}）")
        d.text(x + 4, y + 1, w - 8, 14, label, size=8, font=UD, color="404040", wrap="none", name="子の係の枠の説明")
        n = len(kids)
        gap = 18
        bw = (w - 8 - gap * (n - 1)) / n
        for i, (a, b) in enumerate(kids):
            bx = x + 4 + i * (bw + gap)
            box(d, bx, y + 17, bw, 42, [[(a, {"bold": True, "size": 8 if n == 3 else 9})], [(b, {"size": 8})]],
                fill=BLUE, line=BLUE_LINE, name=f"子の係（{st}・{i + 1}）")
            if i < n - 1:
                d.text(bx + bw, y + 30, gap, 16, sep, size=11, font=UD, color="404040", align="ctr", wrap="none",
                       name="起こす順の印")
        return frame


WAVE = {"prep": ("組1", "組 1 ／ 先に済む物: なし"), "read": ("組2", "組 2 ／ prep の後"),
        "build": ("組3〜4", "組3 共通 → 組4 画面ごと"), "link": ("組5", "組 5 ／ build 全部の後")}
STAGES = [("prep", "最初・モック変更時だけ 1 枚"), ("read", "区切りで 1 枚"),
          ("build", "画面ごとに 1 枚＋共通 1 枚"), ("link", "区切りで 1 枚")]
KIDS = {"prep": None,
        "read": [[("子のサブエージェント", {"bold": True})], "dna-extractor（モックの読み取り）"],
        "build": [[("子のサブエージェント", {"bold": True})], "implementer ⇄ visual-verifier", "（合うまで最大 5 周）"],
        "link": [[("子のサブエージェント", {"bold": True})], "dna-extractor → implementer", "→ visual-verifier"]}
OUTS = {"prep": [[("画面フォルダ", {"bold": True})], "mock.digest.json（モックの値）"],
        "read": [[("画面の読み取りの書類", {"bold": True})], "agent-os/specs/<画面>/"],
        "build": [[("画面のコード", {"bold": True})], "src/"],
        "link": [[("画面どうしの仮のつなぎ", {"bold": True})], "src/prototype/（移動・引き継ぎ）"]}


def build():
    d = Drawing(tag="m2f", name="m2fの動き")
    d.text(26, Y0 + 14, 1600, 56,
           [[("m2f（モックから画面を作る運用）の動き", {})]],
           size=28, font=None, color="000000", wrap="none", name="見出し")
    d.text(620, Y0 + 26, 900, 30, "※ 区切り＝いっしょに作る画面の集まり。区切りごとに (2) 〜 (5) を繰り返す",
           size=16, font=None, color="000000", wrap="none", name="見出しの注")
    d.shape(34, Y0 + 62, FRAME_W, FRAME_H, "rect", fill=None, line=("tx2",), lw=2.25, name="枠")

    # ---- コマンドの並び（人が打つ）----
    CY = Y0 + 100
    # コマンドの帯の幅は、その下の範囲（境目の縦の点線の間）にそろえる
    BOUNDS = [330, 461, 860, 1330, 2202, FRAME_W + 26]
    names = [("/gsd-new-project", "案件を始める"), ("/gsd-discuss-phase N", "画面の表を決める"),
             ("/gsd-plan-phase N", "計画書を書く"), ("/gsd-execute-phase N", "画面を作る"),
             ("/gsd-verify-work N", "受け入れ確認")]
    spec = [(BOUNDS[i] + (0 if i == 0 else 6), BOUNDS[i + 1] - BOUNDS[i] - (6 if i == 0 else 12)) + names[i]
            for i in range(5)]
    cmds = [d.shape(x, CY, w, 60, "homePlate", fill=GREEN, line=GREEN_LINE, lw=0.75, adj=18000,
                    text=[[(a, {"bold": True})], b], size=11, font=YU, align="ctr", ins=(2, 0, 9, 0),
                    color="000000", name=a) for x, w, a, b in spec]
    d.text(60, Y0 + 140, 160, 24, "人（入力するコマンド）", size=12, font=None, color="000000", wrap="none")
    d.connector((56, Y0 + 164), (FRAME_W + 20, Y0 + 164), color=ARROW, lw=AW, tail=False, name="区切り線")
    d.text(60, Y0 + 164, 80, 24, "裏側の処理", size=12, font=None, color="000000", wrap="none")
    MY = Y0 + 205
    d.shape(330, MY, 2072, 41, "roundRect", fill=GREEN, line=GREEN_LINE, adj=8000,
            text="メインエージェント", size=14, font=UD, name="メインエージェント")
    MB = MY + 41
    xs = [c.site(2)[0] for c in cmds]
    for c, x, n in zip(cmds, xs, ["(1)", "(2)", "(3)", "(4)", "(5)"]):
        command(d, (c, 2), (x, MY))
        number(d, x, MY - 46, n)
    # 次の区切りへ（verify-work の上 → discuss の上）
    vx, dxx = cmds[4].x + 40, xs[1]
    with d.group("次の区切りへ"):
        d.connector((vx, CY), (vx, CY - 14), color=NEXT, lw=AW, tail=False, name="次へ")
        d.connector((vx, CY - 14), (dxx, CY - 14), color=NEXT, lw=AW, tail=False, name="次へ")
        d.connector((dxx, CY - 14), (dxx, CY), color=NEXT, lw=AW, name="次へ")
        d.text(cmds[4].x - 120, CY - 32, 150, 16, "次の区切りへ", size=9, font=UD, color="404040", wrap="none",
               align="ctr", name="次の区切りへの説明")

    AY = MY + 90
    # 人が案件に置く物

    # (1) new-project → gsd-roadmapper → 区切りの表
    road = box(d, xs[0] - 55, AY, 110, 58, [[("サブエージェント", {"bold": True})], "gsd-roadmapper"], name="gsd-roadmapper")
    spawn(d, (xs[0], MB), (road, 0))
    roadmap = box(d, xs[0] - 55, AY + 100, 110, 58, [[("区切りの表", {"bold": True})], "ROADMAP.md"], name="ROADMAP")
    write(d, (road, 2), (roadmap, 0))

    # (2) discuss → 区切りの画面の表
    # メインが 2 つの決まりを読んで、2 つの書類を書く
    rule_g = rule_doc(d, 470, AY - 6, 104, 56, [[("GSD の手順書", {"bold": True})], "discuss-phase の手順"],
                      GREEN, name="GSD の手順書")
    rule_k = rule_doc(d, 664, AY - 6, 112, 56, [[("キットの CLAUDE.md", {"bold": True})],
                                                [("画面の表は SCREENS.md へ", {"color": "7030A0", "size": 8})]],
                      BLUE, name="キットの CLAUDE.md")
    context = box(d, 540, AY + 100, 124, 58, [[("区切りの決め事", {"bold": True})], "CONTEXT.md",
                                             [("表は SCREENS.md を指す", {"size": 8})]], name="CONTEXT")
    screens = box(d, 732, AY + 100, 124, 58, [[("区切りの画面の表", {"bold": True})], "SCREENS.md"],
                  fill=BLUE, line=BLUE_LINE, name="SCREENS")
    # メインは決まりを読み（灰の点線）、書類を書く（緑。決まりの箱は通さない）
    for rule, doc in ((rule_g, context), (rule_k, screens)):   # 左で読み、右で書く（書く矢印は箱の上の真ん中へ）
        read(d, (rule.cx, MB), (rule, 0))
        write(d, (doc.cx, MB), (doc, 0))

    # (3) plan → gsd-planner（計画の手引きを読む）→ 計画書。計画書の形は門①が検査する
    PX = 960   # plan の範囲の左寄り（右は計画書のカード）
    planner = box(d, PX - 58, AY, 116, 58, [[("サブエージェント", {"bold": True})], "gsd-planner"], name="gsd-planner")
    spawn(d, (PX, MB), (planner, 0))
    skill = rule_doc(d, PX - 70, AY + 98, 140, 76,
                     [[("計画の手引き", {"bold": True})], "sdd-planner（SKILL.md）",
                      [("config.json の agent_skills", {"color": "7030A0", "size": 8})],
                      [("で、GSD が持たせる", {"color": "7030A0", "size": 8})]], BLUE, name="sdd-planner")
    read(d, (planner, 2), (skill, 0))
    gate1 = gate(d, 884, AY + 204, 160, 54,
                 [[("書いた計画を全部確かめる", {"bold": True, "size": 9})], "門①  gsd-plan-check"], name="門①")
    read(d, (planner, 1), (screens, 3))
    d.text(884, AY + 260, 160, 28, ["Capability sdd-kit が", "plan:post に差し込む"], size=9, font=UD, color="7030A0", wrap="none",
           name="門①の説明")

    CX, CW = 1140, 170
    d.text(CX, AY - 40, 170, 30, "計画書は工程ごとに書く", size=10, font=UD, color="000000", name="計画書の説明")
    cards = []
    for k, (st, rule) in enumerate(STAGES):
        lines = [[(f"計画書 PLAN.md（{st}）", {"bold": True})], [(f"agent_hint: sdd-relay-{st}", {"color": PURPLE})], rule,
                 [(WAVE[st][1], {"bold": True, "color": "C55A11"})]]
        if st == "build":
            cards.append(stack(d, CX, AY + 82 * k - 4, CW, 68, lines, name=f"計画書（{st}）"))
        else:
            cards.append(box(d, CX, AY + 82 * k - 4, CW, 68, lines, name=f"計画書（{st}）"))
    with d.group("順番の決まり方"):
        d.shape(884, AY + 322, 420, 118, "roundRect", fill="FFFFFF", line="C55A11", lw=1.25, adj=6000, align="l",
                anchor="t", text=[[("順番はこう決まる", {"bold": True, "color": "C55A11", "size": 11}),
                                   ("　組（wave）＝いっしょに流す計画書のまとまり", {"color": "C55A11"})],
                                  "① gsd-planner が計画書の頭に、組と先に済む物（depends_on）を書く",
                                  "② execute は組の小さい順に計画書を流し、agent_hint の係を起こす",
                                  "　（同じ組は同時に）",
                                  "③ 組が済むたびに門②が確かめ、キットの案内で流せる計画書を",
                                  "　取り直して、次の組へ"],
                size=10, font=UD, ins=(6, 4, 4, 2), name="順番の決まり方")
    # 門①は区切りの計画書を全部確かめる（点線の枠でひとまとめにし、矢印は枠に刺す）
    plans_box = d.shape(CX - 10, AY - 14, CW + 30, 82 * 3 + 88, "rect", fill=None, line=("tx1", 35, 65), lw=1,
                        dash="sysDash", name="計画書のまとまり")
    for c in cards:
        write(d, (planner, 3), (c, 1))

    # (4) execute → 中継ぎ係（agent_hint の名で起きる）→ 子の係 → 書く物
    RX, RW = 1372, 184
    KX, KW = 1600, 326
    FX, FW = 1970, 190
    relays, kids = [], {}
    for k, (st, _) in enumerate(STAGES):
        y = AY + 82 * k
        role = "（道具を流すだけ）" if st == "prep" else "（検査とコミット。作業は子へ）"
        lines = [[("方法論独自サブエージェント", {"bold": True})], [(f"sdd-relay-{st}", {"color": PURPLE})], role]
        r = (stack if st == "build" else box)(d, RX, y, RW, 58, lines, fill=BLUE, line=BLUE_LINE, name=f"中継ぎ係（{st}）")
        read(d, (r, 1), (cards[k], 3))
        d.shape(r.x, r.y - 17, 112, 15, "roundRect", fill="1F3864", line=None, adj=30000,
                text=[[(f"{WAVE[st][0]}  ・  図2 {st}", {"bold": True})]], size=9, color="FFFFFF", font=YU, ins=(2, 0, 2, 0),
                name=f"対応の札（{st}）")
        out = box(d, FX, y + 4, FW, 50, OUTS[st], fill=BLUE, line=BLUE_LINE, name=f"書く物（{st}）")
        if KIDS[st]:
            kid = kid_group(d, KX, y - 4, KW, 64, st)
            spawn(d, (r, 3), (kid, 1))
            write(d, (kid, 3), (out, 1))
            kids[st] = kid
        else:
            write(d, (r, 3), (out, 1))
            mock = box(d, KX, AY - 40, KW, 30, [[("人が置いたモック", {"bold": True}), ("  mock.svg", {})]],
                       fill=BLUE, line=BLUE_LINE, name="モック")
            read(d, (r, 0), (mock, 1))
            d.text(KX - 12, y + 33, 210, 30, ["子は起こさない：SVG を値と画像に変える",
                                              "決まった変換で、判断が要らないため"], size=9, font=UD, color="595959",
                   name="子なし")
        relays.append(r)
    verifier = box(d, KX, AY + 82 * 4 + 10, KW, 58, [[("サブエージェント", {"bold": True})], "gsd-verifier",
                                                      "（区切りの目標を満たしたか）"], name="gsd-verifier")
    d.shape(KX, verifier.y - 17, 112, 15, "roundRect", fill="1F3864", line=None, adj=30000,
            text=[[("全部の組の後", {"bold": True})]], size=9, color="FFFFFF", font=YU, ins=(2, 0, 2, 0), name="対応の札（照合）")
    # 門②が見るのは、各方法論の係が計画書ごとに残す記録
    records = box(d, FX, AY + 82 * 4 + 6, FW, 46, [[("計画書ごとの記録", {"bold": True})],
                                                   [("SUMMARY.md・RESULT.json", {"size": 8})],
                                                   [("CHECKS.json・deferred-items.md", {"size": 8})]],
                  fill=BLUE, line=BLUE_LINE, name="計画書ごとの記録")
    write(d, (FX + FW / 2, AY + 82 * 3 + 66), (records, 0))
    d.text(FX + FW / 2 + 4, AY + 82 * 3 + 68, 120, 14, "各方法論の係が書く", size=8, font=UD, color="404040", wrap="none",
           name="記録の説明")
    gate2 = gate(d, FX, AY + 82 * 4 + 76, FW + 10, 50, [[("済んだ組の結果を確かめる", {"bold": True})], "門②  gsd-wave-check"],
                 name="門②")
    read(d, (gate1, 3), (plans_box, 1))
    d.text(FX, AY + 82 * 4 + 128, 190, 28, ["Capability sdd-kit が", "execute:wave:post に差し込む"], size=9, font=UD, color="7030A0", wrap="none",
           name="門②の説明")
    read(d, (gate2, 0), (records, 2))
    lane = RX - 22
    exe = box(d, RX, AY + 82 * 4 + 64, RW, 28, [[("gsd-executor（GSD の既定）", {"bold": True})]], name="使わない gsd-executor")
    with d.group("起こす矢印"):
        d.connector((lane, MB), (lane, exe.cy), color=ARROW, lw=AW, tail=False, name="起こす矢印の幹")
        for t in relays + [verifier, exe]:
            spawn(d, (lane, t.cy), (t, 1), name="起こす矢印の枝")
    d.cross(lane + 11, exe.cy, size=20)
    d.text(RX, AY + 82 * 4 - 14, 196, 30, [[("config.json の agent_hint_routing: true", {"color": "7030A0"})],
                                          [("で、agent_hint の係が代わりに起きる", {"color": "7030A0"})]],
           size=8, font=UD, color="7030A0", name="gsd-executor の説明")

    # (5) verify-work → 人が動く画面を確かめる → 足りなければ直しの計画書
    uat = box(d, xs[4] - 80, AY, 160, 58, [[("受け入れ確認", {"bold": True})], "人が動く画面を見て答える"],
              name="受け入れ確認")
    command(d, (xs[4], MB), (uat, 0))
    fix = box(d, xs[4] - 80, AY + 100, 160, 58, [[("足りない所があれば", {"bold": True})], "直す計画書を作る（--gaps）"],
              name="直しの計画書")
    write(d, (uat, 2), (fix, 0))

    d.text(RX, AY - 44, 200, 16, [[("上から順に起きる（組 1 → 組 5）↓", {"bold": True})]], size=10, font=UD,
           color="C55A11", wrap="none", name="起きる順")
    # 図2 で詳しく描く所（execute の中の、工程ごとの方法論の係）を点線の枠で囲む
    fx0, fy0, fx1, fy1 = RX - 8, AY - 48, FX + FW + 10, AY + 82 * 3 + 66
    d.shape(fx0, fy0, fx1 - fx0, fy1 - fy0, "rect", fill=None, line="1F3864", lw=1.25, dash="sysDot",
            name="図2 の枠")
    d.text(fx1 - 222, fy0 + 2, 218, 18, "この枠の中を、図2 で工程ごとに詳しく ↓", size=10, font=UD, color="1F3864",
           wrap="none", align="r", name="図2 の枠の説明")
    # コマンドごとの境目（縦の点線）
    with d.group("コマンドごとの境目"):
        for bx in BOUNDS[1:-1]:
            d.connector((bx, Y0 + 68), (bx, MY - 2), color=("tx1", 35, 65), lw=1, dash="sysDash", tail=False, name="境目")
            d.connector((bx, MB + 2), (bx, Y0 + 62 + FRAME_H - 6), color=("tx1", 35, 65), lw=1, dash="sysDash",
                        tail=False, name="境目")
    legend(d, 60, MY + 60)
    return d, dict(exe=exe, row4=AY + 82 * 4 + 10, map_from=(fx1, fy1), screens=screens, relays=relays, kids=kids, frame_bottom=Y0 + 62 + FRAME_H)


def callouts(d, p):
    FB = p["frame_bottom"]
    with d.group("要点1"):
        d.callout(330, FB + 30, 330, 92,
                  ["要点1", "", "m2f には、人が書類を承認して止まる場面が無い。区切りごとに、話し合いで画面の表を書き、(2) 〜 (5) を繰り返す。"],
                  to=(p["screens"].x + 8, p["screens"].y + p["screens"].h), name="吹き出し")
    with d.group("要点2"):
        r = p["relays"][3]
        d.callout(1339, FB + 30, 260, 120,
                  ["要点2", "", "GSD は計画書の agent_hint を見て、既定の gsd-executor の代わりに、方法論独自のサブエージェントを起こす。"],
                  to=(p["exe"].cx, p["exe"].y + p["exe"].h), name="吹き出し")
    with d.group("要点3"):
        k = p["kids"]["build"]
        d.callout(1763, p["frame_bottom"] + 30, 370, 120,
                  ["要点3", "", "実装の係はモックの画像を開かない（見張りが止める）。モックは mock.digest.json の値で読み、visual-verifier が撮った画面とモックを比べて、合うまで直す。"],
                  to=(k.x + k.w - 4, k.y + k.h), name="吹き出し")


def blank_base(src, out):
    """見本の xlsx から図とセルの文字を外し、シート名を変えた土台を作る。"""
    with zipfile.ZipFile(src) as z, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as o:
        for info in z.infolist():
            data = z.read(info.filename)
            if info.filename == "xl/drawings/drawing1.xml":
                s = re.sub(r"<xdr:(absoluteAnchor|oneCellAnchor|twoCellAnchor)\b.*?</xdr:\1>", "",
                           data.decode("utf-8"), flags=re.S)
                data = s.encode("utf-8")
            elif info.filename == "xl/drawings/_rels/drawing1.xml.rels":
                data = re.sub(r"<Relationship [^>]*/>", "", data.decode("utf-8")).encode("utf-8")
            elif info.filename == "xl/worksheets/sheet1.xml":
                s = re.sub(r"<sheetData>.*?</sheetData>", "<sheetData/>", data.decode("utf-8"), flags=re.S)
                s = re.sub(r"<mergeCells.*?</mergeCells>", "", s, flags=re.S)
                data = s.encode("utf-8")
            elif info.filename == "xl/workbook.xml":
                s = re.sub(r'(<sheet [^>]*name=")[^"]*(")', r"\1m2fの動き\2", data.decode("utf-8"), count=1)
                data = s.encode("utf-8")
            o.writestr(info, data)


if __name__ == "__main__":
    src, out = sys.argv[1], sys.argv[2]
    grid = measure_grid(src)
    base = out + ".base.xlsx"
    blank_base(src, base)
    d, parts = build()
    callouts(d, parts)
    from build_m2f_io import build_io
    d2 = Drawing(tag="m2f-io", name="m2fの入出力")
    y2 = parts["frame_bottom"] + 170
    build_io(d2, y2)
    # 図1 の点線の枠 → 図2 の枠の上（対応の線。枠の右を通す）
    mx, my = parts["map_from"]
    lane_x = 2437
    with d.group("図1 と図2 の対応"):
        d.connector((mx, my - 30), (lane_x, my - 30), color="1F3864", lw=1.25, dash="sysDot", tail=False, name="対応")
        d.connector((lane_x, my - 30), (lane_x, y2 + 100), color="1F3864", lw=1.25, dash="sysDot", tail=False, name="対応")
        d.connector((lane_x, y2 + 100), (1826, y2 + 100), color="1F3864", lw=1.25, dash="sysDot", name="対応")
    write_into(base, out, [d, d2], grid=grid)
    os.remove(base)
    print("書いた", out)
