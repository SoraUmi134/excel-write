"""「GSDに独自処理を挟む2つの方法.xlsx」にパターン 3 の図を足す（書き直す）台本。

使い方:
  python3 build_pattern3.py 元.xlsx 出力.xlsx

パターン 1・2（人が描いた図）の描き方に合わせる:
- 箱は角丸。GSD の物は薄い緑、方法論独自の物は薄い青。方法論で決める値は赤字
- 矢印は橙。箱から箱への矢印は、箱のつなぎ目につなぐ
- 手順は図の中に文章で書かず、矢印の脇の番号（①…）で示す
- 使わない係には ✕
- 問題点は黄色の吹き出しに赤字。該当の箱へ線を出し、参考リンクの札（該当の一文へ飛ぶリンク）を付ける
- パターン全体を 1 つのまとまりにし、吹き出しと札、係と追加のプロンプトは小さなまとまりにする
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "skills", "excel-figure", "scripts"))
from xdraw import Drawing, write_into  # noqa: E402

# ---- 色（パターン 1・2 から読み取った物） ----
GREEN = ("accent6", 20, 80)      # GSD の物・コマンド・メイン
GREEN_LINE = ("accent3",)
BLUE = ("accent1", 20, 80)       # 方法論独自の物
BLUE_LINE = "2F6DB5"
ARROW = ("accent2",)             # 橙の矢印
AW = 1.5                         # 矢印の太さ（テーマの線の型 2 番 = 19050 EMU）
RED = "FF0000"
LINK = ("accent4",)              # 参考リンクの札の青
UD = "BIZ UDGothic"
YU = "Yu Gothic"

DY = 700  # パターン 2 から下へずらす量（pt）

# ---- 参考リンク（該当の一文へ飛ぶ） ----
GH = "https://github.com/open-gsd/gsd-core"
L = {
    "4139": GH + "/blob/main/docs/adr/4139-compact-content-seam.md#:~:text=prose%20dispatch%20is%20not%20guaranteed",
    "step275": GH + "/blob/main/gsd-core/workflows/execute-phase.md#:~:text=Contribution%20dispatch",
    "1689": GH + "/issues/1689#issuecomment-5143711828",
    "planmd": GH + "/blob/main/docs/reference/plan-md.md#per-plan-executor-routing",
    "tools": GH + "/blob/main/docs/FEATURES.md#:~:text=no%20capability%20can%20grant%20tools%20to%20one",
    "firstparty": GH + "/blob/main/docs/explanation/capability-overlay-model.md#first-party-always-wins",
    "ship": GH + "/blob/main/docs/how-to/develop-a-capability.md#ship-skills--and-know-what-you-are-shipping",
    "4740": GH + "/issues/4740",
    "1143": GH + "/blob/main/docs/adr/1143-claude-orchestration-capability.md#:~:text=the%20orchestrator%20IS%20the%20host",
}


def two_line(d, x, y, w, h, head, name_runs, fill, line, head_bold=True, size=11):
    """パターン 1・2 の係の箱: 1 行目に種類、2 行目以降に名前。"""
    return d.shape(x, y, w, h, "roundRect", fill=fill, line=line, lw=0.75, adj=8000,
                   text=[[(head, {"bold": head_bold})]] + name_runs, size=size, font=YU,
                   ins=(4, 2, 4, 2), name=head)


def number(d, x, y, label):
    d.text(x, y, 44, 30, label, size=18, font=None, color="000000", name="番号")


def link_tags(d, x, y, keys):
    for i, k in enumerate(keys):
        d.text(x + i * 82, y, 78, 20, f"参考リンク{i + 1}", size=11, font=UD, color=LINK,
               underline=True, link=L[k], wrap="none", name=f"参考リンク{i + 1}")


def problem(d, x, y, w, h, title, body, to, keys):
    with d.group(title):
        d.callout(x, y, w, h, [title, ""] + body.split("\n"), to=to)
        link_tags(d, x + 8, y + h - 26, keys)


def build():
    d = Drawing(tag="pattern3", name="パターン3")

    # ---- 見出しと枠 ----
    d.text(26.5, 566.9 + DY, 1500, 56,
           "パターン3（Capabilityに全部寄せ、gsd-executorに追加のプロンプトを渡す形）",
           size=28, font=None, color="000000", wrap="none", name="見出し")
    d.shape(33.8, 610.4 + DY, 1517.2, 450.1, "rect", fill=None, line=("tx2",), lw=2.25, name="枠")

    # ---- コマンドの並び（パターン 2 と同じ横位置） ----
    cmds = []
    for x, w, a, b in [(398.2, 156.5, "/gsd-new-project", "案件を始める"),
                       (582.4, 156.2, "/gsd-discuss-phase", "決め事を相談"),
                       (764.7, 156.8, "/gsd-plan-phase", "計画書を書く"),
                       (947.5, 545.2, "/gsd-execute-phase", "作る")]:
        cmds.append(d.shape(x, 696.0 + DY, w, 60.7, "homePlate", fill=GREEN, line=GREEN_LINE, lw=0.75,
                            adj=18000, text=[[(a, {"bold": True})], b], size=11, font=YU,
                            ins=(2, 0, 9, 0), color="000000", name=a))
    d.text(78.1, 755.7 + DY, 150, 24, "人（入力するコマンド）", size=12, font=None, color="000000", wrap="none")
    d.connector((56.1, 778.5 + DY), (1517.6, 778.5 + DY), color=ARROW, lw=AW, tail=False, name="区切り線")
    d.text(78.1, 778.5 + DY, 80, 24, "裏側の処理", size=12, font=None, color="000000", wrap="none")

    MY, MH = 819.0 + DY, 40.9   # メインのバー
    main = d.shape(397.9, MY, 1068.6, MH, "roundRect", fill=GREEN, line=GREEN_LINE, adj=8000,
                   text="メインエージェント", size=14, font=UD, name="メインエージェント")

    AY, AH = 900.0 + DY, 82.0   # 係の段

    # ---- Capability（方法論の中身を全部ここに持つ） ----
    cap = d.shape(70, MY + MH / 2 - 36, 300, 72, "roundRect", fill=BLUE, line=("accent1",), adj=8000,
                  text=[".gsd/capabilities/sdd-kit/capability.json",
                        [('"contributions": [ 計画係へ, 作る係へ ]', {"color": RED})],
                        [('"skills": [ implementer, designer, … ]', {"color": RED})],
                        [('"gates" / "hooks"（今と同じ）', {"color": RED})]],
                  size=10, font=UD, align="l", anchor="ctr", ins=(6, 3, 4, 3), name="Capability")
    d.connector((cap, 3), (main, 1), color=ARROW, lw=AW, name="Capability→メイン")

    # ---- ① ② 案件を始める → 道のり案内 ----
    rx = cmds[0].site(2)[0]
    d.connector((cmds[0], 2), (rx, MY), color=ARROW, lw=AW)
    number(d, rx - 16, 775 + DY, "①")
    road = two_line(d, rx - 52.7, AY, 105.4, AH, "サブエージェント",
                    ["gsd-roadmapper", "（道のり案内）"], GREEN, GREEN_LINE)
    d.connector((rx, MY + MH), (road, 0), color=ARROW, lw=AW)
    number(d, rx - 16, 864 + DY, "②")

    # ---- ③ 相談 ----
    dx = cmds[1].site(2)[0]
    d.connector((cmds[1], 2), (dx, MY), color=ARROW, lw=AW)
    number(d, dx - 16, 777 + DY, "③")

    # ---- ④ ⑤ 計画書を書く → 計画係 ＋ 方法論独自の追加のプロンプト ----
    px = cmds[2].site(2)[0]
    d.connector((cmds[2], 2), (px, MY), color=ARROW, lw=AW)
    number(d, px - 16, 777 + DY, "④")
    with d.group("計画係と追加のプロンプト"):
        planner = two_line(d, px - 52, AY, 104, AH, "サブエージェント",
                           ["gsd-planner", "（計画立案）"], GREEN, GREEN_LINE)
        d.shape(px - 165, AY + 72, 181, 74, "roundRect", fill=BLUE, line=BLUE_LINE, adj=8000,
                text=[[("方法論独自の追加のプロンプト", {"bold": True})],
                      "（Capability の contribution）",
                      [("「PLAN.md に sdd.stage を書け」", {"color": RED})]],
                size=10, font=YU, name="計画係への追加のプロンプト")
    d.connector((px, MY + MH), (planner, 0), color=ARROW, lw=AW)
    number(d, px - 16, 858 + DY, "⑤")

    # ---- ⑥ 計画書 ----
    plan = two_line(d, 930, AY, 130, AH, "計画書 PLAN.md",
                    [[("sdd.stage: build", {"color": RED})], "（agent_hint は無し）"], GREEN, GREEN_LINE)
    d.connector((planner, 3), (plan, 1), color=ARROW, lw=AW)
    number(d, 893, AY + 6, "⑥")

    # ---- ⑦ ⑧ 作る → gsd-executor ＋ 方法論独自の追加のプロンプト ----
    ex = cmds[3].site(2)[0]
    d.connector((cmds[3], 2), (ex, MY), color=ARROW, lw=AW)
    number(d, ex + 4, 777 + DY, "⑦")
    with d.group("作る係と追加のプロンプト"):
        execu = two_line(d, ex - 71.5, AY, 143, AH, "GSDのサブエージェント",
                         ["gsd-executor", "（そのまま使う）"], GREEN, GREEN_LINE)
        d.plus(ex + 98, AY + AH / 2, size=34)
        eprompt = d.shape(ex + 122, AY, 168, AH, "roundRect", fill=BLUE, line=BLUE_LINE, adj=8000,
                          text=[[("方法論独自の追加のプロンプト", {"bold": True})],
                                "（Capability の contribution）",
                                [("「PLAN.md の sdd.stage を読み、", {"color": RED})],
                                [("合う SKILL.md を読んでから作れ」", {"color": RED})]],
                          size=10, font=YU, name="作る係への追加のプロンプト")
    d.connector((ex, MY + MH), (execu, 0), color=ARROW, lw=AW)
    number(d, ex + 4, 858 + DY, "⑧")
    d.connector((execu, 1), (plan, 3), color=ARROW, lw=AW, dash="dash", name="計画書を読む")
    d.text(1068, AY + AH / 2 - 18, 60, 16, "読む", size=10, font=UD, color="000000", wrap="none")

    # ---- ⑨ 手引きを読む ----
    skill = two_line(d, ex - 75, AY + AH + 26, 150, 46, "方法論独自 SKILL.md",
                     ["（Capability の skills）"], BLUE, BLUE_LINE)
    d.connector((execu, 2), (skill, 0), color=ARROW, lw=AW)
    number(d, ex + 4, AY + AH, "⑨")

    # ---- 使わない中継ぎ係 ----
    with d.group("使わない中継ぎ係"):
        relay = two_line(d, ex + 135, AY + AH + 22, 158, 54, "方法論独自サブエージェント",
                         [[("sdd-relay-build", {"color": RED})]], BLUE, BLUE_LINE)
        d.cross(relay.cx, relay.cy, size=56)

    # ---- この形の位置づけ ----
    with d.group("位置づけ"):
        d.text(56, 975 + DY, 330, 52,
               "この形は GSD の想定する使い方ではない。\n作る係の差し替えは agent_hint で行う、と\nGSD が決めている。",
               size=12, font=UD, color="000000")
        link_tags(d, 56, 1030 + DY, ["1689", "4740", "1143"])

    # ---- 問題点 ----
    problem(d, 1580, AY - 30, 370, 132, "問題点1",
            "作る係への追加のプロンプトは、GSD の手順書に「貼れ」と書いてあるだけで、貼る枠が無い。"
            "メインが貼り忘れると届かない。GSD 自身も「文章での差し込みは効く保証が無い」と書いている。",
            to=eprompt.site(3), keys=["4139", "step275"])
    problem(d, 690, 1075 + DY, 270, 160, "問題点2",
            "計画書ごとに作る係を選べない。どの計画書も gsd-executor が作り、作業の切り替えは gsd-executor が"
            " sdd.stage を読んで自分で手引きを選ぶ事に頼る。GSD が agent_hint を作ったのは、ここを文章任せにしないため。",
            to=plan.site(2), keys=["1689", "planmd"])
    problem(d, 975, 1075 + DY, 270, 160, "問題点3",
            "gsd-executor の道具は変えられない。Agent が無いので子のサブエージェントを起こせず、"
            "Playwright も無いので画面を見る作業（visual-verifier）はできない。Capability で道具を足す方法も無い。",
            to=execu.site(1), keys=["tools", "firstparty"])
    problem(d, 1260, 1075 + DY, 290, 160, "問題点4",
            "Capability の skills は、案件だけに入れた場合（--local）は .claude/skills/ に置かれない"
            "（2026-11-05 に 1.20.0 で確認）。gsd-executor には .gsd/capabilities/ の中を直接 Read させる事になる。",
            to=skill.site(2), keys=["ship"])
    return d


if __name__ == "__main__":
    src, dst = sys.argv[1], sys.argv[2]
    removed = write_into(src, dst, build())
    print(f"書きました: {dst}（前のパターン3を {removed} 個消して書き直し）")
