#!/usr/bin/env python3
"""Claude Code の hook（PostToolUse）から呼ぶ点検。

- Write / Edit で Markdown を書いたら、その文書を check_text.py で点検する
- Bash で xlsx を書き出したら、台本で描いた図（descr="xdraw:" を持つ物）を check_figure.py で点検する
問題があれば内容を標準エラーに書き、終了コード 2 で返す（Claude に伝わり、直すまで気づける）。
skill や設定、記憶の Markdown は点検しない（悪い例をわざと書いているため）。
"""
import glob
import json
import os
import subprocess
import sys
import time
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
CHECK_TEXT = os.path.join(HERE, "check_text.py")
CHECK_FIGURE = os.path.join(HERE, "..", "..", "excel-figure", "scripts", "check_figure.py")
SKIP_PARTS = ("/skills/", "/.claude/", "/memory/", "banned_words")
RECENT = 180  # 秒。これより前に書かれた xlsx は点検しない


def run(cmd):
    p = subprocess.run([sys.executable, "-B"] + cmd, capture_output=True, text=True)
    return p.returncode, p.stdout.strip()


def made_by_script(path):
    try:
        with zipfile.ZipFile(path) as z:
            return any('descr="xdraw:' in z.read(n).decode("utf-8", "ignore")
                       for n in z.namelist() if n.startswith("xl/drawings/drawing"))
    except Exception:
        return False


def recent_xlsx(cmd, cwd):
    cands = set()
    for tok in cmd.replace('"', " ").replace("'", " ").split():
        if tok.endswith(".xlsx") and os.path.exists(tok):
            cands.add(tok)
    for base in (cwd, "/private/tmp/claude-501"):
        for p in glob.glob(os.path.join(base, "**", "*.xlsx"), recursive=True):
            cands.add(p)
    now = time.time()
    keep = [p for p in cands if "~$" not in p and "/backup/" not in p]   # 直す前の控えは、わざと古いので見ない
    return sorted(p for p in keep if now - os.path.getmtime(p) < RECENT and made_by_script(p))


def main():
    data = json.load(sys.stdin)
    tool = data.get("tool_name", "")
    inp = data.get("tool_input", {}) or {}
    msgs = []
    if tool in ("Write", "Edit", "MultiEdit"):
        path = inp.get("file_path", "")
        if path.endswith(".md") and not any(k in path for k in SKIP_PARTS) and os.path.exists(path):
            code, out = run([CHECK_TEXT, path])
            if code:
                msgs.append(out)
    elif tool == "Bash" and ".xlsx" in inp.get("command", ""):
        for p in recent_xlsx(inp.get("command", ""), data.get("cwd", os.getcwd())):
            code, out = run([CHECK_FIGURE, p])
            if code:
                lines = out.splitlines()
                more = f"\n  …ほか {len(lines) - 16} 件（全部は check_figure.py で見る）" if len(lines) > 16 else ""
                msgs.append(f"{p}\n" + "\n".join(lines[:16]) + more)
    if msgs:
        sys.stderr.write("渡す前の点検で問題が見つかった。直してから人に渡すこと:\n" + "\n".join(msgs) + "\n")
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
