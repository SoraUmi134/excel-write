# excel-write

Excel の説明図や、人に渡す文書を作るときのやり方と注意点を、Claude Code の skill としてまとめた置き場。
実際に作って、作り直しを頼まれて分かったことを、そのつど書き足していく。

## 中身

| 場所 | 何か |
|---|---|
| `skills/excel-figure/` | Excel のシートにオートシェイプで説明図を描き足す skill。人が描いた図の描き方に合わせる、矢印を箱につなぐ、まとまり（グループ）にする、セルに留める、Excel が無い Mac で見た目を確かめる |
| `skills/excel-figure/scripts/xdraw.py` | 図形を台本で組み立てて xlsx に書き込む部品（python3 の標準の部品だけで動く） |
| `skills/excel-figure/scripts/render_drawing.py` | xlsx の図を PNG / SVG に描いて確かめる道具。図形の一覧も出せる |
| `skills/review-before-delivery/` | 渡す前の点検（ガードレール）の skill。機械の点検・Excel での確かめ・予備知識の無い係に見せる確かめ、指摘を受けたら点検と禁止語に足す決まり、hook から呼ぶ点検 |
| `skills/excel-figure/scripts/check_figure.py` | Excel の図の機械の点検（あふれ・凡例と線の意味・線の向き・横切り・重なり・禁止語・長い段落） |
| `skills/doc-writing/` | 調べた結果を人に渡す文書（資料・根拠リンク集・図の中の短い文）を書くときの決まり |
| `examples/gsd-patterns/` | 実例。`build_pattern3.py` は人が描いたパターン 1・2 に合わせてパターン 3 を足した台本。`build_all.py` は人の了承を得て 3 つのパターンを描き直し、計画書と係を複数描いた台本 |

## Claude Code で使う

skill のフォルダを `~/.claude/skills/` へ繋ぐ（写しではなく繋ぐので、ここを直せばすぐ効く）。

```bash
ln -s ~/Develop/excel-write/skills/excel-figure ~/.claude/skills/excel-figure
```

```bash
ln -s ~/Develop/excel-write/skills/doc-writing ~/.claude/skills/doc-writing
```

```bash
ln -s ~/Develop/excel-write/skills/review-before-delivery ~/.claude/skills/review-before-delivery
```

点検を自動で走らせる hook の入れ方は、`skills/review-before-delivery/SKILL.md` の最後にある。

## 要る物

- python3
- `rsvg-convert`（図を PNG にする。`brew install librsvg`）
- Microsoft Excel（仕上がりを確かめる。無ければ LibreOffice で代わりに見る）

## 書き足し方

作った物が「違う」と言われたら、何が違ったかを skill の「一番大事なこと」か「細かい落とし穴」に足す。
台本で直した所は、実例の台本にも反映する。
