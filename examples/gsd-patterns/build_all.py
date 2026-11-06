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
        d.text(x + 40, y, 120, 18, "起こす", size=10, font=UD, color="000000", wrap="none")
        d.connector((x, y + 29), (x + 34, y + 29), color=GRAY, lw=1.25, dash="dash", name="凡例の点線")
        d.text(x + 40, y + 20, 200, 18, "計画書を読んで従う", size=10, font=UD, color="000000", wrap="none")


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
    d.text(x + header_dx, AY - 42, 120, 30, "工程ごとの計画書\n（画面の数だけある）", size=10, font=UD,
           color="000000", name="計画書の説明")
    cards = []
    for k, st in enumerate(STAGES):
        cards.append(stack(d, x, AY + 80 * k, w, 58, card_lines(st), name=f"計画書（{st}）"))
    for c in cards:
        arrow(d, (planner, 3), (c, 1))
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
    box(d, px - 165, AY + 74, 181, 74,
        [[("sdd-planner", {"color": RED, "bold": True}), ("（SKILL.md）", {"bold": True})],
         [("「PLAN.md の agent_hint に、", {"color": PURPLE})],
         [("工程の中継ぎ係の名前を書け」", {"color": PURPLE})]],
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
    d.text(1120, AY + 262, 300, 30, "agent_hint があるので\ngsd-executor は起きない", size=10, font=UD, color="000000")
    arrow_legend(d, 1300, Y0 + 600)
    with d.group("位置づけ"):
        d.text(64, Y0 + 585, 340, 26, "この方法論自体はGSDとして提供されているもの。", size=14, font=UD, color="000000", wrap="none")
        link_tags(d, 64, Y0 + 615, ["p1_1", "p1_2", "p1_3", "p1_4"])
    d.text(430, Y0 + 585, 470, 40,
           "計画書ごとに agent_hint の係が起きる。工程が変わると、起きる中継ぎ係も変わる"
           "（read → sdd-relay-read、vrt → sdd-relay-vrt）。", size=11, font=UD, color="000000",
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
    d.text(64, Y0 + 585, 320, 40, "どの工程の計画書でも、起きるのは\n同じ gsd-executor。", size=11, font=UD,
           color="000000", name="この形の要点")
    return d


# ---- パターン 3 ------------------------------------------------------------
P3_FRAME_H = 650              # 追加のプロンプトの箱を起こす矢印の通り道に置くため、P1・P2 より縦に長い


def pattern3(Y0):
    d = Drawing(tag="pattern3", name="パターン3")
    cmds, main, MY, AY = common(d, Y0, "パターン3（Capabilityに全部寄せ、gsd-executorに追加のプロンプトを渡す形）",
                                frame_h=P3_FRAME_H)
    RY = AY + 100               # 計画書から右の段は、追加のプロンプトの箱のぶん下げる
    LANE = 1066                 # 起こす矢印の幹の通り道
    cap = box(d, 70, MY - 15.5, 300, 72,
              [".gsd/capabilities/sdd-kit/capability.json",
               [('"contributions": [ 計画係へ, 作る係へ ]', {"color": RED})],
               [('"skills": [ read 用, build 用, vrt 用, … ]', {"color": RED})],
               [('"gates" / "hooks"（今と同じ）', {"color": RED})]],
              fill=BLUE, line=("accent1",), size=10, name="Capability", align="l")
    arrow(d, (cap, 3), (main, 1))
    planner, px = front_steps(d, cmds, MY, AY, ["①", "②", "③", "④", "⑤"],
                              [[("サブエージェント", {"bold": True})], "gsd-planner", "（計画立案）"])
    box(d, px - 165, AY + 74, 181, 74,
        [[("方法論独自の追加のプロンプト", {"bold": True})], "（Capability の contribution）",
         [("「PLAN.md に sdd.stage を書け」", {"color": RED})]],
        fill=BLUE, line=BLUE_LINE, name="計画係への追加のプロンプト")
    cards = plan_column(d, 905, 140, RY, planner,
                        lambda st: [[(f"計画書 PLAN.md（{st}）", {"bold": True})],
                                    [(f"sdd.stage: {st}", {"color": RED})], "（agent_hint は無し）"],
                        "⑥", num_xy=(917, RY - 64), header_dx=36)
    ex = cmds[3].site(2)[0]
    arrow(d, (cmds[3], 2), (ex, MY))
    number(d, ex, MY - 46, "⑦")
    # 起こす矢印の通り道に、作る係への追加のプロンプトを置く（3 体すべての起こす文面に貼られる）
    prompt = box(d, LANE - 150, MY + 60, 300, 70,
                 [[("方法論独自の追加のプロンプト（作る係へ）", {"bold": True})],
                  "（Capability の contribution）",
                  [("「PLAN.md の sdd.stage を読み、", {"color": RED})],
                  [("合う SKILL.md を読んでから作れ」", {"color": RED})]],
                 fill=BLUE, line=BLUE_LINE, name="作る係への追加のプロンプト")
    arrow(d, (LANE, MY + 41), (prompt, 0))
    number(d, LANE - 24, MY + 40, "⑧")
    d.text(LANE + 160, MY + 52, 300, 30,
           "メインが、3 体すべての起こす文面に\nこの文を貼ってから起こす",
           size=10, font=UD, color="000000", name="追加のプロンプトの説明")
    exes, skills = [], []
    for k, st in enumerate(STAGES):
        e = stack(d, 1080, RY + 80 * k, 140, 58,
                  [[("GSDのサブエージェント", {"bold": True})], f"gsd-executor（{NUM[k]}）",
                   [("＋追加のプロンプト", {"color": RED})]], name=f"gsd-executor（{NUM[k]}）")
        s = box(d, 1250, RY + 80 * k + 4, 145, 50,
                [[("方法論独自 SKILL.md", {"bold": True})], f"{st} 用", "（Capability の skills）"],
                fill=BLUE, line=BLUE_LINE, name=f"SKILL.md（{st}）")
        reads(d, cards[k], e)
        arrow(d, (e, 3), (s, 1))
        exes.append(e)
        skills.append(s)
    spawn_bus(d, LANE, prompt.y + prompt.h, exes)
    number(d, 1236, RY - 2, "⑨")
    d.text(1240, RY - 40, 320, 18, "gsd-executor が sdd.stage を読み、自分で選んで読む", size=10, font=UD,
           color="000000", wrap="none", name="手引きの説明")
    with d.group("位置づけ"):
        d.text(56, Y0 + 520, 330, 52,
               "この形は GSD の想定する使い方ではない。\n作る係の差し替えは agent_hint で行う、と\nGSD が決めている。",
               size=12, font=UD, color="000000")
        link_tags(d, 56, Y0 + 575, ["1689", "4740", "1143"])
    arrow_legend(d, 330, Y0 + 660)
    CB = Y0 + 62 + P3_FRAME_H + 26     # 下の吹き出しの上端
    problem(d, 1580, prompt.cy - 75, 380, 150, "問題点1",
            ["作る係への追加のプロンプトは、GSD の手順書に「貼れ」と書いてあるだけで、貼る枠が無い。"
             "メインが貼り忘れると届かない。GSD 自身も「文章での差し込みは効く保証が無い」と書いている。"],
            to=prompt.site(3), keys=["4139", "step275"])
    problem(d, 840, CB, 290, 160, "問題点2",
            ["計画書ごとに作る係を選べない。read・build・vrt のどの計画書も、同じ gsd-executor が作る。"
             "作業の切り替えは、gsd-executor が sdd.stage を読んで自分で手引きを選ぶ事に頼る。"
             "GSD が agent_hint を作ったのは、ここを文章任せにしないため。"],
            to=(exes[2].x + exes[2].w - 20, exes[2].y + exes[2].h), keys=["1689", "planmd"])
    problem(d, 1150, CB, 270, 160, "問題点3",
            ["vrt の計画書でも gsd-executor が作るが、gsd-executor の道具は変えられない。"
             "Playwright が無いので画面を見る作業ができず、Agent が無いので子のサブエージェントも起こせない。"],
            to=(exes[2].x + exes[2].w, exes[2].cy), keys=["tools", "firstparty"])
    problem(d, 1440, CB, 300, 160, "問題点4",
            ["Capability の skills は、案件だけに入れた場合（--local）は .claude/skills/ に置かれない"
             "（2026-11-05 に 1.20.0 で確認）。gsd-executor には .gsd/capabilities/ の中を直接 Read させる事になる。"],
            to=skills[2].site(2), keys=["ship"])
    return d


# ---- パターン 3-B ----------------------------------------------------------
P3B_FRAME_H = 880
RED_DASH = "FF0000"


def badge(d, cx, cy, n, size=22):
    """赤い丸の番号。流れの中の「できない所」と、下の一覧表の行を結ぶ。"""
    return d.shape(cx - size / 2, cy - size / 2, size, size, "ellipse", fill=RED_DASH, line=None,
                   text=[[(str(n), {"bold": True, "color": "FFFFFF"})]], size=12, font=UD,
                   ins=(0, 0, 0, 0), name=f"できない所 {n}")


def ghost(d, x, y, w, h, lines, name):
    """本当は要るが、この形ではできない物（灰の点線の箱）。"""
    return d.shape(x, y, w, h, "roundRect", fill="F2F2F2", line="7F7F7F", lw=1, dash="dash", adj=8000,
                   text=lines, size=10, font=YU, color="595959", ins=(4, 2, 4, 2), name=name)


def cannot(d, a, b, n):
    """本当は要る線（赤の点線）に ✕ と番号を付ける。"""
    (x1, y1) = a[0].site(a[1]) if hasattr(a[0], "site") else a
    (x2, y2) = b[0].site(b[1]) if hasattr(b[0], "site") else b
    d.connector(a, b, color=RED_DASH, lw=1.5, dash="dash", name="できない線")
    mx, my = (x1 + x2) / 2, (y1 + y2) / 2
    d.cross(mx, my, size=22, color=RED_DASH)
    badge(d, mx, my - 20, n)


def table(d, x, y, cols, rows, row_h=34):
    """図形で描く一覧表。cols=[(見出し, 幅)]、rows=[[セル, ...]]。セルは 文字列 か (文字列, 色)。
    1 列目の先頭の数字は赤い丸の番号にする。"""
    with d.group("一覧表"):
        cx = x
        for title, w in cols:
            d.shape(cx, y, w, row_h, "rect", fill=("bg2",), line="A6A6A6", lw=0.75,
                    text=[[(title, {"bold": True})]], size=11, font=UD, color="000000", name="表の見出し")
            cx += w
        for r, row in enumerate(rows):
            ry = y + row_h * (r + 1)
            cx = x
            for c, ((title, w), cell) in enumerate(zip(cols, row)):
                if isinstance(cell, dict):          # 参考リンクの札の並び
                    d.shape(cx, ry, w, row_h, "rect", fill="FFFFFF", line="A6A6A6", lw=0.75, name="表のセル")
                    for k, key in enumerate(cell["links"]):
                        link_tag(d, cx + 6 + k * 82, ry + 7, k + 1, key)
                else:
                    text, color = cell if isinstance(cell, tuple) else (cell, "222222")
                    d.shape(cx, ry, w, row_h, "rect", fill="FFFFFF", line="A6A6A6", lw=0.75,
                            text=text, size=10, font=UD, color=color, align="l", ins=(30 if c == 0 else 6, 2, 4, 2),
                            name="表のセル")
                    if c == 0:
                        badge(d, cx + 14, ry + row_h / 2, r + 1, size=20)
                cx += w


def pattern3b(Y0):
    d = Drawing(tag="pattern3b", name="パターン3-B")
    cmds, main, MY, AY = common(d, Y0, "パターン3-B（Capabilityに全部寄せると、何ができないか）",
                                frame_h=P3B_FRAME_H)
    RY = AY + 100
    LANE = 1066
    cap = box(d, 70, MY - 15.5, 300, 72,
              [".gsd/capabilities/sdd-kit/capability.json",
               [('"contributions": [ 計画係へ, 作る係へ ]', {"color": RED})],
               [('"skills": [ read 用, build 用, vrt 用, … ]', {"color": RED})],
               [('"gates" / "hooks"（今と同じ）', {"color": RED})]],
              fill=BLUE, line=("accent1",), size=10, name="Capability", align="l")
    arrow(d, (cap, 3), (main, 1))
    badge(d, cap.x + cap.w - 12, cap.y + 46, 5)
    planner, px = front_steps(d, cmds, MY, AY, ["①", "②", "③", "④", "⑤"],
                              [[("サブエージェント", {"bold": True})], "gsd-planner", "（計画立案）"])
    box(d, px - 165, AY + 74, 181, 74,
        [[("方法論独自の追加のプロンプト", {"bold": True})], "（Capability の contribution）",
         [("「PLAN.md に sdd.stage を書け」", {"color": RED})]],
        fill=BLUE, line=BLUE_LINE, name="計画係への追加のプロンプト")
    cards = plan_column(d, 905, 140, RY, planner,
                        lambda st: [[(f"計画書 PLAN.md（{st}）", {"bold": True})],
                                    [(f"sdd.stage: {st}", {"color": RED})], "（agent_hint は無し）"],
                        "⑥", num_xy=(917, RY - 64), header_dx=36)
    ex = cmds[3].site(2)[0]
    arrow(d, (cmds[3], 2), (ex, MY))
    number(d, ex, MY - 46, "⑦")
    prompt = box(d, LANE - 150, MY + 60, 300, 70,
                 [[("方法論独自の追加のプロンプト（作る係へ）", {"bold": True})],
                  "（Capability の contribution）",
                  [("「PLAN.md の sdd.stage を読み、", {"color": RED})],
                  [("合う SKILL.md を読んでから作れ」", {"color": RED})]],
                 fill=BLUE, line=BLUE_LINE, name="作る係への追加のプロンプト")
    arrow(d, (LANE, MY + 41), (prompt, 0))
    number(d, LANE - 24, MY + 40, "⑧")
    badge(d, prompt.x + prompt.w + 14, prompt.y + 14, 2)
    d.text(prompt.x + prompt.w + 30, prompt.y + 4, 280, 30,
           "貼る枠が無い。3 体に届くかはメインの読み方次第", size=10, font=UD, color=RED_DASH,
           name="できない所 2 の説明")
    exes = []
    for k, st in enumerate(STAGES):
        e = stack(d, 1080, RY + 80 * k, 140, 58,
                  [[("GSDのサブエージェント", {"bold": True})], f"gsd-executor（{NUM[k]}）",
                   [("＋追加のプロンプト", {"color": RED})]], name=f"gsd-executor（{NUM[k]}）")
        reads(d, cards[k], e)
        exes.append(e)
    spawn_bus(d, LANE, prompt.y + prompt.h, exes)
    badge(d, LANE + 14, RY - 26, 1)
    d.text(LANE + 28, RY - 35, 300, 18, "どの計画書でも同じ gsd-executor（係を選ぶ口が無い）",
           size=10, font=UD, color=RED_DASH, wrap="none", name="できない所 1 の説明")
    # 本当は要る物（パターン 1 で中継ぎ係が子の係にやらせていること）
    wants = {
        "read": ["本当は要る: 子の係 dna-extractor に", "モックを読ませ、監査役と直しを回す"],
        "build": ["本当は要る: implementer と", "visual-verifier を最大 5 往復"],
        "vrt": ["本当は要る: Playwright で画面を撮り、", "基準の画像と比べる"],
    }
    d.text(1300, RY - 15, 260, 16, "今の方法論で中継ぎ係がやっていること", size=10, font=UD,
           color="595959", wrap="none", name="本当は要る物の説明")
    for k, st in enumerate(STAGES):
        g = ghost(d, 1300, RY + 80 * k + 2, 240, 54, wants[st], name=f"本当は要る物（{st}）")
        cannot(d, (exes[k], 3), (g, 1), 4 if st == "vrt" else 3)
    with d.group("位置づけ"):
        d.text(56, Y0 + 520, 330, 52,
               "この形は GSD の想定する使い方ではない。\n作る係の差し替えは agent_hint で行う、と\nGSD が決めている。",
               size=12, font=UD, color="000000")
        link_tags(d, 56, Y0 + 575, ["1689", "4740", "1143"])
    with d.group("凡例"):
        arrow_legend(d, 330, Y0 + 520)
        d.connector((330, Y0 + 569), (364, Y0 + 569), color=RED_DASH, lw=1.5, dash="dash", name="凡例の赤線")
        d.text(370, Y0 + 560, 280, 18, "本当は要るが、この形ではできない", size=10, font=UD, color=RED_DASH, wrap="none")
    # 一覧表: 何ができないから、Capability だけには寄せられないのか
    TY = RY + 270
    d.text(64, TY - 30, 900, 24, "Capability に全部寄せたときに、できること・できないこと（赤い番号は上の図の場所）",
           size=14, font=UD, color="000000", wrap="none", name="一覧表の見出し")
    cols = [("方法論に要ること", 330), ("今の方法論（パターン1）", 330), ("Capability だけ（パターン3）", 470), ("根拠", 260)]
    rows = [
        ["計画書ごとに、作る係を選ぶ", "○ 計画書の agent_hint で中継ぎ係を選ぶ",
         ("✕ 選ぶ口が無い。どの計画書も gsd-executor が作る", RED_DASH), {"links": ["1689", "planmd"]}],
        ["方法論の手順を、作る係に確実に渡す", "○ 中継ぎ係のファイルそのものが手順",
         ("△ 起こす文面に貼るだけ。貼る枠が無く、届く保証も無い", "C55A11"), {"links": ["4139", "step275"]}],
        ["子の係を起こして、往復させる", "○ 中継ぎ係が Agent の道具で子を起こす",
         ("✕ gsd-executor に Agent の道具が無く、足す方法も無い", RED_DASH), {"links": ["tools", "firstparty"]}],
        ["画面を見て確かめる（Playwright）", "○ visual-verifier などの子が道具を持つ",
         ("✕ gsd-executor に Playwright が無く、足す方法も無い", RED_DASH), {"links": ["tools"]}],
        ["手引き（SKILL.md）を係の手元に置く", "○ キットの導入（gsd-setup）が置く",
         ("△ 案件だけの導入（--local）では置かれない", "C55A11"), {"links": ["ship"]}],
    ]
    table(d, 64, TY, cols, rows)
    concl_y = TY + 34 * (len(rows) + 1) + 10
    d.shape(64, concl_y, 1390, 40, "rect", fill=("accent2", 20, 80), line=("accent2",), lw=1,
            text=[[("結論: ", {"bold": True}),
                   ("✕ の 3 つ（係を選ぶ・子を起こす・画面を見る）は、Capability の文章では埋められない。"
                    "作る係の選び方は agent_hint（GSD が決めた正規の道）を残し、Capability は門・見張り・文章の差し込みに使う。", {})]],
            size=12, font=UD, color="000000", align="l", ins=(10, 2, 8, 2), name="結論")
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
    Y_P3B = Y_P3 + 62 + P3_FRAME_H + 26 + 160 + 60
    drawings = [pattern1(Y_P1), pattern2(Y_P2), pattern3(Y_P3), pattern3b(Y_P3B)]
    n = write_into(src, dst, drawings, grid=grid, replace_hand_drawn=True,
                   sheet_edit=move_cells(grid, Y_P3B + 62 + P3B_FRAME_H + 20))
    print(f"書きました: {dst}（外した図 {n} 個、描いたパターン {len(drawings)} つ）")
