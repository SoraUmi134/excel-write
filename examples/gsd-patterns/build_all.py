"""「GSDに独自処理を挟む…方法.xlsx」のパターン 1・2・3 を、計画書と係を複数描く形で描き直す台本。

使い方:
  python3 build_all.py 元.xlsx 出力.xlsx

- 人が描いたパターン 1・2 の文字・リンク・吹き出しの中身は引き継ぎ、図だけ描き直す（人の了承済み）。
- 計画書は工程ごと（read・build・vrt）に 3 段、それぞれ重なったカードで「画面の数だけある」を示す。
  パターン 1 は計画書ごとに agent_hint で別の中継ぎ係が起き、パターン 2・3 はどの計画書でも gsd-executor が起きる。
- セル BB18 の「./claude/agents/〇〇.md」は図の中の文字に移し、B61 の「以上」は最後の図の下へ動かす。
"""
import os
import re
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "skills", "excel-figure", "scripts"))
from xdraw import Drawing, measure_grid, write_into  # noqa: E402

GREEN = ("accent6", 20, 80)
GREEN_LINE = ("accent3",)
BLUE = ("accent1", 20, 80)
BLUE_LINE = "2F6DB5"
ARROW = ("accent2",)
AW = 1.5
GRAY = ("tx1", 50, 50)       # 「計画書を読んで従う」の点線
NUM = ["1 体目", "2 体目", "3 体目"]
RED = "FF0000"
PURPLE = ("accent5",)
LINK = ("accent4",)
UD = "BIZ UDGothic"
YU = "Yu Gothic"
Y_P1, Y_P2, Y_P3 = 0, 720, 1610   # 各パターンの上端（pt）。吹き出しの無い P1 の下は詰める
STAGES = ["read", "build", "vrt"]
CHILDREN = {"read": "dna-extractor（読み取り）",
            "build": "implementer ⇄ visual-verifier",
            "vrt": "vrt-test-implementer"}

GH = "https://github.com/open-gsd/gsd-core"
L = {
    # パターン 1（人の図のまま）
    "p1_1": GH + "/blob/603b2c6/docs/reference/plan-md.md#per-plan-executor-routing",
    "p1_2": GH + "/blob/603b2c6/docs/CONFIGURATION.md#agent-skills-injection",
    "p1_3": GH + "/blob/main/docs/CONFIGURATION.md#workflow-toggles:~:text=workflow.agent_hint_routing",
    "p1_4": "https://docs.opengsd.net/core/configuration/settings#section-breakdown",
    # パターン 2（人の図のまま）
    "p2_2363": GH + "/blob/main/docs/adr/2363-capability-instruction-surface-trust.md#:~:text=Third%2Dparty%20agents%5B%5D%20are%20never%20staged%20into%20the%20agent%27s%20instruction%20context.",
    "p2_manifest": GH + "/blob/main/docs/reference/capability-manifest.md#feature-body-role-feature",
    "p2_ship": GH + "/blob/main/docs/how-to/develop-a-capability.md#ship-skills--and-know-what-you-are-shipping",
    "p2_5738": GH + "/issues/5738#issuecomment-6476728511",
    # パターン 3
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


# ---- 部品 -------------------------------------------------------------------
def box(d, x, y, w, h, lines, fill=GREEN, line=GREEN_LINE, size=10, name="箱", align="ctr"):
    return d.shape(x, y, w, h, "roundRect", fill=fill, line=line, lw=0.75, adj=8000, align=align,
                   text=lines, size=size, font=YU, ins=(6 if align == "l" else 4, 2, 4, 2), name=name)


def stack(d, x, y, w, h, lines, fill=GREEN, line=GREEN_LINE, size=10, name="重なった箱"):
    """同じ物が複数あることを示す、重なったカード。前の 1 枚を返す。"""
    with d.group(name):
        for k in (2, 1):
            d.shape(x + 4 * k, y + 3 * k, w, h, "roundRect", fill=fill, line=line, lw=0.75, adj=8000, name="重なり")
        return box(d, x, y, w, h, lines, fill, line, size, name)


def arrow(d, a, b, **kw):
    kw.setdefault("color", ARROW)
    kw.setdefault("lw", AW)
    d.connector(a, b, **kw)


def reads(d, card, agent):
    """計画書 → 係: 起こすのではなく、係が計画書を読んで従う（灰の点線）。"""
    d.connector((card, 3), (agent, 1), color=GRAY, lw=1.25, dash="dash", name="計画書を読む")


def spawn_bus(d, lane_x, top_y, targets, side=None):
    """メインから係を起こす矢印。メインのバーから幹を下ろし、段ごとに枝を出して係の上へ入れる。
    side=(箱, ✕を付けるか) は、幹から横へ枝を出す箱（使わない係など）。"""
    ys = [t.y - 9 for t in targets]
    bottom = max(ys + ([side[0].cy] if side else []))
    d.connector((lane_x, top_y), (lane_x, bottom), color=ARROW, lw=AW, tail=False, name="起こす矢印の幹")
    for t, y in zip(targets, ys):
        d.connector((lane_x, y), (t.cx, y), color=ARROW, lw=AW, tail=False, name="起こす矢印の枝")
        d.connector((t.cx, y), (t, 0), color=ARROW, lw=AW, name="起こす矢印")
    if side:
        box_, crossed = side
        d.connector((lane_x, box_.cy), (box_, 3), color=ARROW, lw=AW, name="起こす矢印")
        if crossed:
            d.cross((lane_x + box_.x + box_.w) / 2, box_.cy, size=24)


def arrow_legend(d, x, y):
    """矢印の凡例。"""
    with d.group("凡例"):
        d.connector((x, y + 9), (x + 34, y + 9), color=ARROW, lw=AW, name="凡例の矢印")
        d.text(x + 40, y, 340, 18, "メインエージェントがサブエージェントを起こす", size=10, font=UD, color="000000", wrap="none")
        d.connector((x, y + 29), (x + 34, y + 29), color=GRAY, lw=1.25, dash="dash", name="凡例の点線")
        d.text(x + 40, y + 20, 340, 18, "サブエージェントが計画書を読んで、その内容に従う", size=10, font=UD, color="000000", wrap="none")


def number(d, cx, y, label):
    d.text(cx - 22, y, 44, 30, label, size=18, font=None, color="000000", align="ctr", name="番号")


def link_tag(d, x, y, i, key):
    d.text(x, y, 78, 20, f"参考リンク{i}", size=11, font=UD, color=LINK,
           underline=True, link=L[key], wrap="none", name=f"参考リンク{i}")


def link_tags(d, x, y, keys):
    for i, k in enumerate(keys):
        link_tag(d, x + i * 82, y, i + 1, k)


def problem(d, x, y, w, h, title, body, to, keys):
    with d.group(title):
        d.callout(x, y, w, h, [title, ""] + body, to=to)
        if keys:
            link_tags(d, x + 8, y + h - 26, keys)


def common(d, Y0, title, exec_text=("/gsd-execute-phase", "作る"), exec_align="ctr", frame_h=602):
    """見出し・枠・コマンドの並び・人と裏側の区切り・メインのバー。"""
    t = title if isinstance(title, list) else [[(title, {})]]
    d.text(26, Y0 + 16, 1500, 56, t, size=28, font=None, color="000000", wrap="none", name="見出し")
    d.shape(34, Y0 + 62, 1517, frame_h, "rect", fill=None, line=("tx2",), lw=2.25, name="枠")
    CY = Y0 + 82
    cmds = []
    for x, w, a, b in [(398.2, 156.5, "/gsd-new-project", "案件を始める"),
                       (582.4, 156.2, "/gsd-discuss-phase", "決め事を相談"),
                       (764.7, 156.8, "/gsd-plan-phase", "計画書を書く"),
                       (947.5, 545.2) + tuple(exec_text)]:
        al = exec_align if a == "/gsd-execute-phase" else "ctr"
        cmds.append(d.shape(x, CY, w, 60.7, "homePlate", fill=GREEN, line=GREEN_LINE, lw=0.75, adj=18000,
                            text=[[(a, {"bold": True})], b], size=11, font=YU, align=al,
                            ins=(12 if al == "l" else 2, 0, 9, 0),
                            color="000000", name=a))
    d.text(78, Y0 + 140, 150, 24, "人（入力するコマンド）", size=12, font=None, color="000000", wrap="none")
    d.connector((56, Y0 + 164), (1517.6, Y0 + 164), color=ARROW, lw=AW, tail=False, name="区切り線")
    d.text(78, Y0 + 164, 80, 24, "裏側の処理", size=12, font=None, color="000000", wrap="none")
    MY = Y0 + 205
    main = d.shape(397.9, MY, 1068.6, 41, "roundRect", fill=GREEN, line=GREEN_LINE, adj=8000,
                   text="メインエージェント", size=14, font=UD, name="メインエージェント")
    return cmds, main, MY, Y0 + 290


def front_steps(d, cmds, MY, AY, nums, planner_lines):
    """① ② ③ ④ ⑤: 案件を始める → 道のり案内、相談、計画書を書く → 計画係。"""
    rx, dx, px = (cmds[i].site(2)[0] for i in range(3))
    for c, x, n in ((cmds[0], rx, nums[0]), (cmds[1], dx, nums[2]), (cmds[2], px, nums[3])):
        arrow(d, (c, 2), (x, MY))
        number(d, x, MY - 46, n)
    road = box(d, rx - 52.7, AY, 105.4, 82, [[("サブエージェント", {"bold": True})], "gsd-roadmapper", "（道のり案内）"])
    arrow(d, (rx, MY + 41), (road, 0))
    number(d, rx, MY + 42, nums[1])
    planner = box(d, px - 52, AY, 104, 82, planner_lines)
    arrow(d, (px, MY + 41), (planner, 0))
    number(d, px, MY + 42, nums[4])
    return planner, px


def plan_column(d, x, w, AY, planner, card_lines, num=None, num_xy=None, header_dx=12):
    """計画書のカード 3 段（工程ごと）。計画係から 3 本の矢印。num_xy で番号の置き場を変えられる。"""
    d.text(x + header_dx, AY - 42, 130, 30, "計画書は工程ごとに、\n画面 1 つにつき 1 枚", size=10, font=UD,
           color="000000", name="計画書の説明")
    cards = []
    for k, st in enumerate(STAGES):
        cards.append(stack(d, x, AY + 80 * k, w, 58, card_lines(st), name=f"計画書（{st}）"))
    for c in cards:
        arrow(d, (planner, 3), (c, 1))
    d.text(x, AY + 222, 172, 26, "read: 見本の読み取り　build: 画面作り\nvrt: 画面を撮って見本と比べる",
           size=9, font=UD, color="595959", name="工程の意味")
    if num:
        number(d, *(num_xy or (x - 12, AY - 32)), num)
    return cards


# ---- パターン 1 ------------------------------------------------------------
def pattern1(Y0):
    d = Drawing(tag="pattern1", name="パターン1")
    cmds, main, MY, AY = common(d, Y0, [[("パターン1（現状の方法論の形）", {}),
                                        ("※ただし、現状の方法論でも独自hookはパターン2を利用している", {"size": 20})]])
    box(d, 90, MY + 20, 285, 81,
        [".planning/config.json", "", '"agent_skills": {',
         [('    "gsd-planner":  [".claude/skills/sdd-planner"]', {"color": RED})], "  }"],
        fill=BLUE, line=("accent1",), size=10, name="config.json", align="l")
    planner, px = front_steps(d, cmds, MY, AY, ["(1)", "(2)", "(3)", "(4)", "(5)"],
                              [[("サブエージェント", {"bold": True})], [("gsd-planner", {"color": RED})], "（計画立案）"])
    box(d, px - 220, AY + 74, 236, 80,
        [[("sdd-planner", {"color": RED, "bold": True}), ("（SKILL.md）", {"bold": True})],
         [("「PLAN.md の agent_hint に、その工程の", {"color": PURPLE})],
         [("方法論独自サブエージェントの名前を書け」", {"color": PURPLE})]],
        fill=BLUE, line=BLUE_LINE, name="sdd-planner の手引き")
    cards = plan_column(d, 925, 150, AY, planner,
                        lambda st: [[(f"計画書 PLAN.md（{st}）", {"bold": True})],
                                    [("agent_hint:", {"color": PURPLE})],
                                    [(f"sdd-relay-{st}", {"color": PURPLE})]],
                        "(6)", (900, AY + 10))
    ex = cmds[3].site(2)[0]
    arrow(d, (cmds[3], 2), (ex, MY))
    number(d, ex, MY - 46, "(7)")
    d.text(1145, AY - 40, 400, 18, "方法論独自サブエージェント（.claude/agents/〇〇.md に置く）",
           size=10, font=UD, color="000000", wrap="none", name="中継ぎ係の説明")
    relays = []
    for k, st in enumerate(STAGES):
        relay = stack(d, 1120, AY + 80 * k, 170, 58,
                      [[("方法論独自サブエージェント", {"bold": True})], [(f"sdd-relay-{st}", {"color": PURPLE})],
                       "（計画書 1 枚に 1 体）"], fill=BLUE, line=BLUE_LINE, name=f"中継ぎ係（{st}）")
        child = box(d, 1330, AY + 80 * k + 4, 190, 50,
                    [[("子のサブエージェント", {"bold": True})], CHILDREN[st]], fill=BLUE, line=BLUE_LINE,
                    name=f"子の係（{st}）")
        reads(d, cards[k], relay)
        arrow(d, (relay, 3), (child, 1))
        relays.append(relay)
    exe = box(d, 925, AY + 250, 150, 52, [[("GSDのサブエージェント", {"bold": True})], "gsd-executor"],
              name="使わない gsd-executor")
    spawn_bus(d, 1101, MY + 41, relays, side=(exe, True))
    number(d, 1101 + 18, MY + 42, "(8)")
    number(d, 1310, AY - 2, "(9)")
    d.text(1120, AY + 250, 330, 44, "計画書に agent_hint があるので、GSD 既定の\ngsd-executor は起きず、agent_hint に書いた\nサブエージェントが代わりに起きる", size=10, font=UD, color="000000")
    arrow_legend(d, 1300, Y0 + 600)
    with d.group("位置づけ"):
        d.text(64, Y0 + 585, 340, 26, "この方法論自体はGSDとして提供されているもの。", size=14, font=UD, color="000000", wrap="none")
        link_tags(d, 64, Y0 + 615, ["p1_1", "p1_2", "p1_3", "p1_4"])
    d.text(430, Y0 + 585, 470, 40,
           "GSD は、計画書の agent_hint に書かれたサブエージェントを起こす。工程ごとに agent_hint が違うので、"
           "工程に合ったサブエージェントが起きる（read の計画書なら sdd-relay-read）。", size=11, font=UD, color="000000",
           name="この形の要点")
    return d


# ---- パターン 2 ------------------------------------------------------------
def pattern2(Y0):
    d = Drawing(tag="pattern2", name="パターン2")
    cmds, main, MY, AY = common(d, Y0, "パターン2（Capabilityのfeature経由での機能追加）", exec_align="l")
    planner, px = front_steps(d, cmds, MY, AY, ["①", "②", "③", "④", "⑤"],
                              [[("サブエージェント", {"bold": True})], "gsd-planner", "（計画立案）"])
    cards = plan_column(d, 905, 140, AY, planner,
                        lambda st: [[(f"計画書 PLAN.md（{st}）", {"bold": True})], "（agent_hint は無し）"])
    ex = cmds[3].site(2)[0]
    arrow(d, (cmds[3], 2), (ex, MY))
    number(d, ex, MY - 46, "⑥")
    exes = []
    for k, st in enumerate(STAGES):
        e = stack(d, 1080, AY + 80 * k, 140, 58,
                  [[("GSDのサブエージェント", {"bold": True})], f"gsd-executor（{NUM[k]}）"],
                  name=f"gsd-executor（{NUM[k]}）")
        reads(d, cards[k], e)
        exes.append(e)
    spawn_bus(d, 1066, MY + 41, exes)
    arrow_legend(d, 420, Y0 + 610)
    with d.group("使えない中継ぎ係"):
        relay = box(d, 1240, AY, 160, 58, [[("方法論独自サブエージェント", {"bold": True})], "sdd-relay-build"],
                    fill=BLUE, line=BLUE_LINE)
        d.cross(relay.cx, relay.cy, size=56)
    skill = box(d, 1410, AY, 138, 58, [[("方法論独自SKILL.md", {"bold": True})], "「〇〇をして」"],
                fill=BLUE, line=BLUE_LINE)
    for target, n in ((relay, "⑦"), (skill, "⑧")):
        cx = target.cx
        d.text(cx - 54, Y0 + 112, 108, 22, "post（またはpre）", size=11, font=None, color="000000", align="ctr", wrap="none")
        arrow(d, (cx, Y0 + 142.7), (target, 0))
        number(d, cx + 26, MY + 42, n)
    problem(d, 880, Y0 + 690, 250, 110, "問題点1",
            ["GSDのサブエージェントが必ず動いてしまう。gsd-executorは実装なども行うため、勝手に欲しくもないコードが作られてしまう。",
             "", "こっちは、従わせられない方が問題。"],
            to=(exes[2].cx + 20, exes[2].y + exes[2].h), keys=[])
    with d.group("問題点2"):
        d.callout(1170, Y0 + 650, 316, 205,
                  ["問題点2", "",
                   "GSDのCapabilityにagents（サブエージェント）を定義しても、GSDがその定義ファイルを、.claude/agents/にコピーしてくれず、Claude Codeがサブエージェントを起こせない。",
                   "", "", "",
                   "ただし、2026/10/28にGSDのメンテナーが「修正する。ちゃんとGSDのagentsに置かれたものは、.claude/agents/に置かれるようにする」と表明。だが、まだ修正は未済。"],
                  to=relay.site(2))
        link_tag(d, 1183, Y0 + 650 + 92, 1, "p2_2363")
        link_tag(d, 1265, Y0 + 650 + 92, 2, "p2_manifest")
        link_tag(d, 1347, Y0 + 650 + 92, 3, "p2_ship")
        link_tag(d, 1183, Y0 + 650 + 180, 4, "p2_5738")
    problem(d, 1580, AY - 10, 400, 112, "問題点3",
            ["独自SKILL.mdをGSDに定義すれば、それはClaude Codeで利用できるようになるが、あくまで独自SKILL.mdを呼ぶことができるのはメインエージェントのみ。",
             "", "（そもそも現時点ではできないが）独自サブエージェントを立ち上げたとしても、そのサブエージェントに独自のSKILL.mdを読ませることができない。"],
            to=skill.site(3), keys=[])
    d.text(64, Y0 + 570, 330, 50, "計画書に agent_hint が無いので、GSD は\nどの工程の計画書でも、既定の gsd-executor を\n起こす。", size=11, font=UD,
           color="000000", name="この形の要点")
    return d


# ---- パターン 3 ------------------------------------------------------------
# 文は 1 項目 1〜2 行。流れ・✕・番号・表で伝え、説明の長文は図に入れない（人の指摘。2026-11-06）。
P3_FRAME_H = 650
P3_CALLOUT_H = 130

CAP_LINES = [".gsd/capabilities/sdd-kit/capability.json",
             [('"contributions": 依頼文に書き足す文', {"color": RED})],
             [('"skills": 方法論の手引き SKILL.md', {"color": RED})],
             [('"gates" / "hooks": 検査と見張り（今も使用中）', {"color": RED})]]

PLANNER_ADD = [[("gsd-planner への依頼文に書き足す文", {"bold": True})],
               [("「計画書に工程（sdd.stage）を書け」", {"color": RED})]]

EXECUTOR_ADD = [[("gsd-executor への依頼文に書き足す文", {"bold": True})],
                [("「sdd.stage を読み、その工程の", {"color": RED})],
                [("手引きを読んでから作業せよ」", {"color": RED})]]

WHY_P3 = ("考えた理由: 独自サブエージェントを案件に置くキット（方法論の部品一式）の手順は、GSD の版上げのたびに直す手間がかかる。"
          "Capability に全部載せられれば、この手間が無くなる。")


def capability_box(d, MY, main):
    cap = box(d, 70, MY + 48, 300, 62, CAP_LINES, fill=BLUE, line=("accent1",), size=10,
              name="Capability", align="l")
    arrow(d, (cap, 0), (main, 1))
    return cap


def executor_column(d, cards, RY):
    exes = []
    for k, st in enumerate(STAGES):
        e = stack(d, 1080, RY + 80 * k, 140, 58,
                  [[("GSDのサブエージェント", {"bold": True})], f"gsd-executor（{NUM[k]}）",
                   [("＋書き足された文", {"color": RED})]], name=f"gsd-executor（{NUM[k]}）")
        reads(d, cards[k], e)
        exes.append(e)
    return exes


def p3_common(d, Y0, title, frame_h, RY_off):
    """パターン 3 と 3-B で同じ部分（コマンド・メイン・Capability・計画係・計画書・書き足す文・gsd-executor）。"""
    cmds, main, MY, AY = common(d, Y0, title, frame_h=frame_h)
    RY = AY + RY_off
    LANE = 1066
    cap = capability_box(d, MY, main)
    planner, px = front_steps(d, cmds, MY, AY, ["①", "②", "③", "④", "⑤"],
                              [[("サブエージェント", {"bold": True})], "gsd-planner", "（計画立案）"])
    box(d, px - 205, AY + 74, 221, 44, PLANNER_ADD, fill=BLUE, line=BLUE_LINE, name="gsd-planner への書き足し")
    cards = plan_column(d, 905, 140, RY, planner,
                        lambda st: [[(f"計画書 PLAN.md（{st}）", {"bold": True})],
                                    [(f"sdd.stage: {st}", {"color": RED})], "（agent_hint は無し）"],
                        "⑥", num_xy=(893, RY - 64), header_dx=36)
    ex = cmds[3].site(2)[0]
    arrow(d, (cmds[3], 2), (ex, MY))
    number(d, ex, MY - 46, "⑦")
    prompt = box(d, LANE - 140, MY + 58, 280, 54, EXECUTOR_ADD, fill=BLUE, line=BLUE_LINE,
                 name="gsd-executor への書き足し")
    arrow(d, (LANE, MY + 41), (prompt, 0))
    number(d, LANE - 24, MY + 40, "⑧")
    exes = executor_column(d, cards, RY)
    spawn_bus(d, LANE, prompt.y + prompt.h, exes)
    return cmds, main, MY, AY, RY, LANE, cap, prompt, cards, exes


def pattern3(Y0):
    d = Drawing(tag="pattern3", name="パターン3")
    _, _, MY, AY, RY, LANE, cap, prompt, cards, exes = p3_common(
        d, Y0, "パターン3（Capabilityに全部寄せ、gsd-executorに追加のプロンプトを渡す形）", P3_FRAME_H, 100)
    d.text(52, Y0 + 70, 342, 40, WHY_P3, size=10, font=UD, color="000000", name="考えた理由")
    d.text(prompt.x + prompt.w + 12, prompt.y - 16, 300, 30,
           "メインが、起こす 3 体すべての依頼文に\n同じ文を書き足す", size=10, font=UD, color="000000", name="書き足しの説明")
    skills = []
    for k, st in enumerate(STAGES):
        s = box(d, 1250, RY + 80 * k + 4, 145, 50,
                [[("方法論の手引き SKILL.md", {"bold": True})], f"{st} の工程用"],
                fill=BLUE, line=BLUE_LINE, name=f"SKILL.md（{st}）")
        arrow(d, (exes[k], 3), (s, 1))
        skills.append(s)
    number(d, 1236, RY - 2, "⑨")
    d.text(1240, RY - 30, 320, 18, "gsd-executor が自分で選んで読む", size=10, font=UD, color="000000",
           wrap="none", name="手引きの説明")
    with d.group("位置づけ"):
        d.text(56, Y0 + 540, 340, 30, "GSD が正式に用意した agent_hint を使わない、\nGSD の想定外の使い方", size=11, font=UD, color="000000")
        link_tags(d, 56, Y0 + 576, ["1689", "4740", "1143"])
    arrow_legend(d, 56, Y0 + 610)
    CB = Y0 + 62 + P3_FRAME_H + 26
    problem(d, 1580, prompt.cy - 60, 360, 120, "問題点2",
            ["書き足しは、メインが手順書を読んで自分で入れる仕組みで、入る保証が無い。"
             "入らないと、gsd-executor は方法論の手順を知らずに作る。"],
            to=prompt.site(3), keys=["4139", "step275"])
    problem(d, 690, CB, 320, P3_CALLOUT_H, "問題点1",
            ["read・build・vrt のどの計画書でも gsd-executor が起きる。"
             "工程に必要な道具と手順を持つサブエージェントに任せられない。"],
            to=(exes[2].x + exes[2].w - 20, exes[2].y + exes[2].h), keys=["1689", "planmd"])
    problem(d, 1030, CB, 310, P3_CALLOUT_H, "問題点3",
            ["gsd-executor はサブエージェントを起こせない（Agent ツールが無く、Capability でも足せない）。作った本人が点検するしかない。"
             "Playwright MCP も無く、見た目の確かめ方は作り直しが要る。"],
            to=(exes[2].x + exes[2].w, exes[2].cy), keys=["tools", "firstparty"])
    problem(d, 1360, CB, 320, P3_CALLOUT_H, "問題点4",
            ["GSD を案件ごとに入れると、Capability の手引き（SKILL.md）は .claude/skills/ に写されず、"
             "サブエージェントが自動で読み込まない（1.20.0 で確認）。置き場所を依頼文で教える必要がある。"],
            to=skills[2].site(2), keys=["ship"])
    return d


# ---- パターン 3-B ----------------------------------------------------------
RED_DASH = "FF0000"
ORANGE = "C55A11"


def badge(d, cx, cy, n, size=22):
    """赤い丸の番号。流れの中の「できない所」と、下の一覧表の行を結ぶ。"""
    return d.shape(cx - size / 2, cy - size / 2, size, size, "ellipse", fill=RED_DASH, line=None,
                   text=[[(str(n), {"bold": True, "color": "FFFFFF"})]], size=12, font=UD,
                   ins=(0, 0, 0, 0), name=f"できない所 {n}")


def ghost(d, x, y, w, h, lines, name):
    """今の方法論では必要だが、この形ではできない物（灰の点線の箱）。"""
    return d.shape(x, y, w, h, "roundRect", fill="F2F2F2", line="7F7F7F", lw=1, dash="dash", adj=8000,
                   text=lines, size=10, font=YU, color="595959", ins=(4, 2, 4, 2), name=name)


def cannot(d, a, b, n, weak=False):
    """今の方法論では必要な線。できない所は赤の点線に ✕、確実でない所はオレンジの点線（✕ なし）。"""
    (x1, y1) = a[0].site(a[1])
    (x2, y2) = b[0].site(b[1])
    d.connector(a, b, color=ORANGE if weak else RED_DASH, lw=1.5, dash="dash", name="できない線")
    mx, my = (x1 + x2) / 2, (y1 + y2) / 2
    if not weak:
        d.cross(mx, my, size=22, color=RED_DASH)
    badge(d, mx, my - 20, n)


def table(d, x, y, cols, rows, row_h=34, head_h=30):
    """図形で描く一覧表。cols=[(見出し, 幅)]、rows=[[セル, ...]]。セルは 文字列 か (文字列, 色) か {"links": [...]}。
    1 列目には行の番号を赤い丸で付ける。"""
    with d.group("一覧表"):
        cx = x
        for title, w in cols:
            d.shape(cx, y, w, head_h, "rect", fill=("bg2",), line="A6A6A6", lw=0.75,
                    text=[[(title, {"bold": True})]], size=11, font=UD, color="000000", name="表の見出し")
            cx += w
        for r, row in enumerate(rows):
            ry = y + head_h + row_h * r
            cx = x
            for c, ((title, w), cell) in enumerate(zip(cols, row)):
                if isinstance(cell, dict):
                    d.shape(cx, ry, w, row_h, "rect", fill="FFFFFF", line="A6A6A6", lw=0.75, name="表のセル")
                    for k, key in enumerate(cell["links"]):
                        link_tag(d, cx + 4, ry + 2 + k * 18, k + 1, key)
                else:
                    text, color = cell if isinstance(cell, tuple) else (cell, "222222")
                    d.shape(cx, ry, w, row_h, "rect", fill="FFFFFF", line="A6A6A6", lw=0.75,
                            text=text, size=10, font=UD, color=color, align="l", anchor="ctr",
                            ins=(30 if c == 0 else 6, 2, 4, 2), name="表のセル")
                    if c == 0:
                        badge(d, cx + 14, ry + row_h / 2, r + 1, size=20)
                cx += w
        return y + head_h + row_h * len(rows)


P3B_ROWS = [
    ["工程ごとに合ったサブエージェントを起こす",
     "○ 計画書の agent_hint に書いたサブエージェントを、GSD が起こす",
     ("✕ Capability には、起こすサブエージェントを指定する手段が無い。どの計画書でも gsd-executor が起きる", RED_DASH),
     "工程に必要な道具・手順を持つサブエージェントに任せられない", {"links": ["1689", "planmd"]}],
    ["方法論の手順を、作業するサブエージェントに確実に伝える",
     "○ 独自サブエージェントのファイルに手順が書いてあり、起きた時点で必ず持っている",
     ("△ メインエージェントが依頼文に書き足すのを当てにするしかなく、入る保証が無い", ORANGE),
     "入らないと、gsd-executor は方法論の手順を知らずに作る", {"links": ["4139", "step275"]}],
    ["別のサブエージェントに点検させる",
     "○ 独自サブエージェントが Agent ツール（サブエージェントを起こす道具）で子を起こし、点検させる",
     ("✕ gsd-executor はサブエージェントを起こせない（Agent ツールを持っておらず、Capability でも足せない）", RED_DASH),
     "作った本人が点検するしかなく、見落としが残る", {"links": ["tools", "firstparty"]}],
    ["画面を開いて見た目を確かめる",
     "○ 子のサブエージェント（visual-verifier）が Playwright MCP で画面を操作しながら見る",
     ("△ gsd-executor には Playwright MCP が無く、Capability でも足せない。Bash（コマンドの実行）で撮影スクリプトを動かせば撮れる", ORANGE),
     "見た目の確かめ方を、撮影スクリプトで撮って比べる形に作り直す必要がある", {"links": ["tools"]}],
    ["手引き（SKILL.md）を、サブエージェントが自動で読む場所に置く",
     "○ キットの導入の手順（gsd-setup）が .claude/skills/ に置く",
     ("△ GSD を案件ごとに入れると、Capability の手引きは .claude/skills/ に写されない（GSD 1.20.0 で確認）", ORANGE),
     "サブエージェントは手引きを自動では読まない。置き場所を依頼文で教えるしかない", {"links": ["ship"]}],
]
P3B_ROW_H = 46
P3B_TABLE_Y = 262           # RY からの表の位置
P3B_FRAME_H = 290 + 150 + P3B_TABLE_Y + 30 + P3B_ROW_H * len(P3B_ROWS) + 90


def pattern3b(Y0):
    d = Drawing(tag="pattern3b", name="パターン3-B")
    _, _, MY, AY, RY, LANE, cap, prompt, cards, exes = p3_common(
        d, Y0, "パターン3-B（パターン3 で、何ができないか）", P3B_FRAME_H, 150)
    badge(d, cap.x + cap.w - 12, cap.y + 34, 5)
    badge(d, prompt.x + prompt.w + 14, prompt.cy, 2)
    badge(d, LANE + 14, prompt.y + prompt.h + 20, 1)
    wants = {
        "read": ["今の方法論: 読み取る役と", "点検する役を別々に起こす"],
        "build": ["今の方法論: 作る役と見た目を", "比べる役を往復させて直す"],
        "vrt": ["今の方法論: Playwright MCP で画面を", "操作しながら撮り、見本の画像と比べる"],
    }
    for k, st in enumerate(STAGES):
        g = ghost(d, 1265, RY + 80 * k + 6, 240, 46, wants[st], name=f"今の方法論では必要な物（{st}）")
        cannot(d, (exes[k], 3), (g, 1), 4 if st == "vrt" else 3, weak=(st == "vrt"))
    with d.group("凡例"):
        arrow_legend(d, 56, RY + 150)
        d.connector((56, RY + 199), (90, RY + 199), color=RED_DASH, lw=1.5, dash="dash", name="凡例の赤線")
        d.text(96, RY + 190, 340, 18, "✕ 今の方法論では必要だが、できない", size=10, font=UD, color=RED_DASH, wrap="none")
        d.connector((56, RY + 219), (90, RY + 219), color=ORANGE, lw=1.5, dash="dash", name="凡例のオレンジ線")
        d.text(96, RY + 210, 340, 18, "△ できるが、確実でない／作り直しが要る", size=10, font=UD, color=ORANGE, wrap="none")
    TY = RY + P3B_TABLE_Y
    cols = [("必要なこと", 270), ("今の方法論（パターン1）", 320), ("Capability だけ（パターン3）", 380),
            ("何が困るか", 330), ("根拠", 100)]
    end_y = table(d, 64, TY, cols, P3B_ROWS, row_h=P3B_ROW_H)
    with d.group("結論"):
        d.shape(64, end_y + 12, 1300, 46, "rect", fill=("accent2", 20, 80), line=("accent2",), lw=1,
                text=[[("結論: ", {"bold": True}),
                       ("✕ の 1・3 は GSD の仕組みで決まっていて、Capability では変えられない。△ は補足として使う程度。", {})],
                      [("→ パターン1（agent_hint でサブエージェントを選ぶ形）のまま進め、Capability は検査と見張りに使う。", {})]],
                size=11, font=UD, color="000000", align="l", anchor="ctr", ins=(10, 2, 8, 2), name="結論")
        d.text(64, end_y + 64, 1300, 18,
               "先の見込み: 「Capability に入れたサブエージェントも GSD が案件に置く」修正が承認済み（2026-10-28、未リリース）。"
               "出れば、サブエージェントの配り方は Capability に寄せられる（選ぶのは agent_hint のまま）。",
               size=10, font=UD, color="000000", name="先の見込み")
        link_tag(d, 1375, end_y + 26, 1, "p2_5738")
    return d


def move_cells(grid, end_y):
    """セル BB18 を空にし、B61 の「以上」を最後の図の下の行へ動かす。"""
    row = int(end_y * 12700 // grid.row_h) + 2

    def edit(sheet):
        sheet = re.sub(r'<row r="18"[^>]*><c r="BB18"[^>]*>.*?</c></row>', "", sheet)
        sheet = re.sub(r'<row r="\d+" spans="2:2"><c r="B\d+" t="s"><v>0</v></c></row>',
                       f'<row r="{row}" spans="2:2"><c r="B{row}" t="s"><v>0</v></c></row>', sheet)
        sheet = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="B{row}"/>', sheet)
        return sheet
    return edit


if __name__ == "__main__":
    src, dst = sys.argv[1], sys.argv[2]
    grid = measure_grid(src)
    Y_P3B = Y_P3 + 62 + P3_FRAME_H + 26 + P3_CALLOUT_H + 60
    drawings = [pattern1(Y_P1), pattern2(Y_P2), pattern3(Y_P3), pattern3b(Y_P3B)]
    n = write_into(src, dst, drawings, grid=grid, replace_hand_drawn=True,
                   sheet_edit=move_cells(grid, Y_P3B + 62 + P3B_FRAME_H + 20))
    print(f"書きました: {dst}（外した図 {n} 個、描いたパターン {len(drawings)} つ）")
