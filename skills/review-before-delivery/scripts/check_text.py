#!/usr/bin/env python3
"""人に渡す文書（Markdown など）を機械的に点検する。

使い方: python3 check_text.py 文書.md [...]
見つける物: 禁止語（../banned_words.txt）、長すぎる 1 文（コードの区切りと表は除く）。
終了コード: 見つかった数が 0 なら 0、あれば 1。
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
BANNED = os.path.join(HERE, "..", "banned_words.txt")
LONG_SENTENCE = 100   # 1 文の長さの上限（字）。リンクの URL は数えない


def load_banned():
    out = []
    for line in open(BANNED, encoding="utf-8"):
        line = line.split("#", 1)[0].strip()
        if line:
            word, _, alt = line.partition("\t")
            out.append((word.strip(), alt.strip()))
    return out


def check(path):
    found = []
    banned = load_banned()
    in_code = False
    for no, line in enumerate(open(path, encoding="utf-8"), 1):
        if line.strip().startswith("```"):
            in_code = not in_code
            continue
        if in_code or line.lstrip().startswith("|"):
            continue
        plain = re.sub(r"\]\([^)]*\)", "]", line)
        plain = re.sub(r"`[^`]*`", "x", plain)
        for word, alt in banned:
            if word in plain:
                found.append(f"{path}:{no} [禁止語] 「{word}」→ {alt or '相手の言葉かふだんの言葉に'}")
        for sent in re.split(r"(?<=[。！？])", plain.strip()):
            if len(sent) > LONG_SENTENCE:
                found.append(f"{path}:{no} [長い文] {len(sent)} 字（上限 {LONG_SENTENCE}）: 「{sent[:30]}…」")
    return found


def main():
    found = [f for p in sys.argv[1:] for f in check(p)]
    for f in found:
        print(f)
    if not found:
        print("点検: 問題は見つからなかった（機械で拾える範囲）")
    return 1 if found else 0


if __name__ == "__main__":
    sys.exit(main())
