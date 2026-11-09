"""m2f の 2 枚目の図: 工程ごとの入力・出力と、方法論の工夫。build_m2f.py から呼ぶ。

線の種類は 3 つ（読む・書く・返す）。凡例も同じ表（LEGEND2）から描く。
"""
from build_m2f import (BLUE, BLUE_LINE, GRAY, UD, WRITE, YU, AW, box, read, write)

FILE_LINE = "2F6DB5"
IMG = ("accent2", 20, 80)          # 画像のファイル（薄い橙）
IMG_LINE = ("accent2",)
TOOL = "D9D9D9"                    # 道具（灰）
BACK = ("accent5",)                # 返す（紫）
NAVY = "1F3864"
RED = "C00000"
COLS = [210, 430, 650, 870, 1090, 1310]   # 段の中の列（左端）
CW = 190
TIP_X, TIP_W = 1525, 290


def file(d, x, y, lines, w=CW, h=34, name="ファイル"):
    return d.shape(x, y, w, h, "foldedCorner", fill="FFFFFF", line=FILE_LINE, lw=0.75, adj=12000,
                   text=lines, size=9, font=YU, ins=(4, 1, 8, 1), name=name)


def image(d, x, y, lines, w=CW, h=34, name="画像のファイル"):
    return d.shape(x, y, w, h, "foldedCorner", fill=IMG, line=IMG_LINE, lw=0.75, adj=12000,
                   text=lines, size=9, font=YU, ins=(4, 1, 8, 1), name=name)


def agent(d, x, y, lines, w=CW, h=46, name="係"):
    return box(d, x, y, w, h, lines, fill=BLUE, line=BLUE_LINE, size=10, name=name)


def tool(d, x, y, lines, w=CW, h=46, name="道具"):
    return d.shape(x, y, w, h, "rect", fill=TOOL, line=("tx1", 50, 50), lw=0.75, text=lines, size=10, font=YU,
                   name=name)


def back(d, pts, name="返す"):
    """返す（紫の実線）。pts は折れ線の点。最後の 1 本だけ矢じり。"""
    for i in range(len(pts) - 1):
        d.connector(pts[i], pts[i + 1], color=BACK, lw=AW, tail=(i == len(pts) - 2), name=name)


def note(d, x, y, text, color=RED, w=CW, name="一行の注"):
    d.text(x, y, w, 16, text, size=9, font=UD, color=color, wrap="none", name=name)


def tip(d, x, y, n, text):
    with d.group(f"工夫{n}"):
        d.shape(x, y, 22, 22, "ellipse", fill=NAVY, line=None, text=[[(str(n), {"bold": True})]], size=11,
                color="FFFFFF", font=YU, ins=(0, 0, 0, 0), name="番号")
        d.text(x + 28, y - 2, TIP_W - 30, 44, text, size=10, font=UD, color="000000", name="工夫の文")


def badge(d, cx, y, n):
    d.shape(cx - 11, y, 22, 22, "ellipse", fill=NAVY, line=None, text=[[(str(n), {"bold": True})]], size=11,
            color="FFFFFF", font=YU, ins=(0, 0, 0, 0), name="番号")


LEGEND2 = [(GRAY, 1.25, "dash", "読む（係・道具が、矢印の先の物を読む）"),
           (WRITE, AW, None, "書く（矢印の先の物を作る）"),
           (BACK, AW, None, "返す（指摘を前の係へ戻す）")]


def legend2(d, x, y):
    with d.group("凡例（2 枚目）"):
        d.text(x, y - 22, 120, 20, [[("凡例", {"bold": True})]], size=11, font=UD, color="000000", wrap="none")
        for k, (color, lw, dash, label) in enumerate(LEGEND2):
            yy = y + 20 * k
            d.connector((x, yy + 9), (x + 34, yy + 9), color=color, lw=lw, dash=dash, name="凡例の線")
            d.text(x + 40, yy, 260, 18, label, size=10, font=UD, color="000000", wrap="none")
        yy = y + 20 * len(LEGEND2) + 6
        rows = [(lambda a, b: file(d, a, b, "", w=30, h=18, name="凡例の色（値）"), "値・文字のファイル"),
                (lambda a, b: image(d, a, b, "", w=30, h=18, name="凡例の色（画像）"), "画像のファイル（PNG）"),
                (lambda a, b: agent(d, a, b, "", w=30, h=18, name="凡例の色（係）"), "方法論のサブエージェント"),
                (lambda a, b: tool(d, a, b, "", w=30, h=18, name="凡例の色（道具）"), "道具（キットのスクリプト）")]
        for k, (mk, label) in enumerate(rows):
            mk(x + 2, yy + 22 * k)
            d.text(x + 40, yy + 22 * k, 260, 18, label, size=10, font=UD, color="000000", wrap="none")


def lane_title(d, y, h, title, sub):
    box(d, 60, y + 8, 128, h - 16, [[(title, {"bold": True, "size": 13})], sub], fill=("bg1", 5, 95),
        line=("tx1", 50, 50), name=f"段の名前（{title}）")


def build_io(d, Y):
    d.text(26, Y + 14, 1700, 40, [[("工程ごとに、係へ何を渡し、何が返るか（方法論の工夫）", {})]],
           size=24, font=None, color="000000", wrap="none", name="見出し（2 枚目）")
    lanes = ["prep", "read", "build", "link"]
    hs = {"prep": 175, "read": 125, "build": 215, "link": 180}
    top = Y + 62
    total = sum(hs.values())
    d.shape(34, top, 1790, total + 20, "rect", fill=None, line=("tx2",), lw=2.25, name="枠（2 枚目）")
    ys, y = {}, top + 10
    for k in lanes:
        ys[k] = y
        y += hs[k]
    for k in lanes[1:]:
        d.connector((44, ys[k]), (1814, ys[k]), color=("tx1", 25, 75), lw=0.75, dash="dash", tail=False,
                    name="区切り線")
    d.connector((TIP_X - 15, top + 6), (TIP_X - 15, top + total + 14), color=("tx1", 25, 75), lw=0.75,
                dash="dash", tail=False, name="区切り線")
    d.text(TIP_X, top + 2, 200, 18, [[("方法論の工夫", {"bold": True})]], size=12, font=UD, color=NAVY,
           wrap="none", name="工夫の見出し")

    # ---- prep: モックを取り込む ----
    y = ys["prep"]
    lane_title(d, y, hs["prep"], "prep", "モックを取り込む")
    svg = file(d, COLS[0], y + 62, [[("mock.svg", {"bold": True})], "人が案件に置くモック"], h=40, name="mock.svg")
    t = tool(d, COLS[1], y + 59, [[("道具 gsd-prep.mjs", {"bold": True})], "svg-digest → svg-rasterize"])
    read(d, (t, 1), (svg, 3))
    OX, OW = COLS[2] + 30, 260
    outs = [file(d, OX, y + 22, [[("mock.digest.json", {"bold": True})], "値の辞書（文言・色・寸法・余白・部品）"],
                 w=OW, name="digest"),
            image(d, OX, y + 60, [[("mock.png", {"bold": True})], "お手本の画像（実装と同じブラウザで描く）"], w=OW, name="mock.png"),
            image(d, OX, y + 98, [[("mock.read-NN.png", {"bold": True})], "分けた画像（長辺 1568px 以下）"], w=OW, name="タイル"),
            image(d, OX, y + 136, [[("mock-assets/*.png", {"bold": True})], "ロゴ・アイコンの部品"], w=OW, name="部品")]
    for o in outs:
        write(d, (t, 3), (o, 1))
    badge(d, OX + OW + 18, y + 28, 1)
    badge(d, OX + OW + 18, y + 66, 2)
    tip(d, TIP_X, y + 26, 1, "SVG を「値の辞書」に直して渡す。数 MB の SVG を、そのまま係に読ませない")
    tip(d, TIP_X, y + 88, 2, "お手本の画像は、作った画面と同じブラウザで描く（描き方の差を消す）")

    # ---- read: モックを読む ----
    y = ys["read"]
    lane_title(d, y, hs["read"], "read", "モックを読む")
    dig = file(d, COLS[0], y + 40, [[("mock.digest.json", {"bold": True})], "構造・文言・部品の節"], name="digest（read）")
    ext = agent(d, COLS[1] + 20, y + 34, [[("dna-extractor", {"bold": True})], "読む（別の回でもう一度確かめる）"],
                name="dna-extractor（read）")
    note(d, COLS[1] + 20, y + 84, "✕ 画像は開かない（見張りが止める）")
    read(d, (ext, 1), (dig, 3))
    dna = file(d, COLS[2] + 60, y + 22, [[("design-dna.md", {"bold": True})], "見た目の決め事"], name="design-dna")
    req = file(d, COLS[2] + 60, y + 64, [[("requirements-simple.md", {"bold": True})], "画面の要件の下書き"], name="req")
    write(d, (ext, 3), (dna, 1))
    write(d, (ext, 3), (req, 1))
    badge(d, COLS[1] + 20 + CW + 18, y + 14, 3)
    tip(d, TIP_X, y + 30, 3, "係は画像を見ずに、値だけで画面を読む（別の係がもう一度読んで確かめる）")

    # ---- build: 画面を作り、お手本と比べる ----
    y = ys["build"]
    lane_title(d, y, hs["build"], "build", "画面を作り\nお手本と比べる")
    ins = [file(d, COLS[0], y + 18, [[("design-dna.md", {"bold": True})], "見た目の決め事"], name="dna（build）"),
           file(d, COLS[0], y + 58, [[("mock.digest.json", {"bold": True})], "値の辞書"], name="digest（build）")]
    imp = agent(d, COLS[1], y + 52, [[("implementer", {"bold": True})], "画面のコードを書く"], name="implementer")
    note(d, COLS[1], y + 102, "✕ 画像は開かない（見張りが止める）")
    for i in ins:
        read(d, (imp, 1), (i, 3))
    src = file(d, COLS[2], y + 58, [[("src/", {"bold": True})], "画面のコード"], name="src")
    write(d, (imp, 3), (src, 1))
    note(d, COLS[2], y + 96, "部品の PNG は、方法論の係が src へ写す", color="595959", w=210)
    note(d, COLS[2], y + 114, "（実装の係はファイル名で使うだけ）", color="595959", w=210)
    ver = agent(d, COLS[3], y + 52, [[("visual-verifier", {"bold": True})], "撮って、お手本と比べる"], name="visual-verifier")
    note(d, COLS[3], y + 34, "画像を見るのは、この係だけ", color="C55A11")
    read(d, (ver, 1), (src, 3))
    mk = image(d, COLS[4] + 20, y + 14, [[("mock.png", {"bold": True})], "お手本の画像"], name="mock.png（build）")
    shot = image(d, COLS[4] + 20, y + 56, [[("screenshot.png", {"bold": True})], "撮った画面（shoot.mjs）"], name="撮った画面")
    with d.group("3 枚組"):
        for k in (2, 1):
            d.shape(COLS[4] + 20 + 4 * k, y + 98 + 3 * k, CW, 40, "foldedCorner", fill=IMG, line=IMG_LINE, lw=0.75,
                    adj=12000, name="重なり")
        tiles = image(d, COLS[4] + 20, y + 98, [[("compare/ の 3 枚組", {"bold": True})], "お手本・作った画面・差分"],
                      h=40, name="3 枚組")
    read(d, (ver, 3), (mk, 1))
    write(d, (ver, 3), (shot, 1))
    write(d, (ver, 3), (tiles, 1))
    back(d, [(ver.cx, ver.y + ver.h + 18), (ver.cx, y + 190), (imp.cx, y + 190), (imp.cx, y + 122)])
    d.connector((ver.cx, ver.y + ver.h), (ver.cx, ver.y + ver.h + 18), color=BACK, lw=AW, tail=False, name="返す")
    d.text(imp.cx + 10, y + 172, 460, 18, "指摘を文字で返す（画面の座標と digest の値）。合うまで最大 5 周",
           size=9, font=UD, color="7030A0", wrap="none", name="返すの説明")
    badge(d, COLS[3] + CW + 12, y + 24, 4)
    badge(d, COLS[4] + 20 + CW + 22, y + 108, 6)
    badge(d, COLS[3] + CW - 6, y + 102, 5)
    tip(d, TIP_X, y + 12, 4, "画像で比べるのは visual-verifier だけ。実装の係は画像を見ない（役目を分ける）")
    tip(d, TIP_X, y + 74, 5, "合否は数字で決める: 違う画素 2% 以下・塊 500px 以下・部品ごと 50% 以下")
    tip(d, TIP_X, y + 136, 6, "大きい画面は長辺 1568px 以下のタイルに分け、3 枚組と拡大で見せる（縮むと細部が潰れるため）")

    # ---- link: 画面どうしを仮につなぐ ----
    y = ys["link"]
    lane_title(d, y, hs["link"], "link", "画面どうしを\n仮につなぐ")
    lin = file(d, COLS[0], y + 40, [[("design-dna.md", {"bold": True})], "＋ 状態モックの digest"], h=40, name="dna（link）")
    ext2 = agent(d, COLS[1], y + 37, [[("dna-extractor", {"bold": True})], "つなぎ表を下書きする"], name="dna-extractor（link）")
    read(d, (ext2, 1), (lin, 3))
    table = file(d, COLS[2], y + 26, [[("つなぎ表", {"bold": True})], "agent-os/prototype/<区切り>.md",
                                       "移動・引き継ぎ・入力チェック・状態"], h=64, name="つなぎ表")
    write(d, (ext2, 3), (table, 1))
    imp2 = agent(d, COLS[3], y + 37, [[("implementer", {"bold": True})], "画面ごとに、仮のつなぎを書く"], name="implementer（link）")
    read(d, (imp2, 1), (table, 3))
    proto = file(d, COLS[4], y + 26, [[("src/prototype/", {"bold": True})], "画面の移動・値の引き継ぎ",
                                       "目印 // PROTOTYPE: 番号"], h=64, name="試作のつなぎ")
    write(d, (imp2, 3), (proto, 1))
    ver2 = agent(d, COLS[5], y + 31, [[("visual-verifier", {"bold": True})], "表の行ごとに押して確かめる",
                                      "（Playwright で、移る・値が出るか）"], h=58, name="visual-verifier（link）")
    read(d, (ver2, 1), (proto, 3))
    d.connector((ver2.cx, ver2.y + ver2.h), (ver2.cx, ver2.y + ver2.h + 24), color=BACK, lw=AW, tail=False, name="返す")
    back(d, [(ver2.cx, ver2.y + ver2.h + 24), (ver2.cx, y + 158), (imp2.cx, y + 158), (imp2.cx, imp2.y + imp2.h)])
    d.text(imp2.cx + 10, y + 140, 300, 18, "足りなければ、その画面のつなぎに戻す。最大 5 周", size=9, font=UD, color="7030A0",
           wrap="none", name="返すの説明")
    badge(d, COLS[4] + CW - 6, y + 14, 7)
    badge(d, COLS[5] + CW - 6, y + 14, 8)
    tip(d, TIP_X, y + 20, 7, "仮のつなぎに目印を付ける。運用 sdd で、この目印を探して本物のつなぎに置き換える")
    tip(d, TIP_X, y + 86, 8, "つなぎ表 1 枚に集め、表の行ごとに押して確かめる。決めきれない所は「要確認」に残す")

    legend2(d, 60, top + total + 60)
    return top + total + 20
