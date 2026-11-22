#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
csvfix.py —— CSV 乱码修复 / 批量合并 / 去重，单文件零依赖。

Excel 双击 CSV 中文乱码，基本都是编码不是 UTF-8-BOM 造成的。
这个脚本把 CSV 统一转成「Excel 友好的 UTF-8-BOM」，同时能合并多份、
自动识别分隔符、按整行去重，默认只预览不写盘。

用法
-----
  python3 csvfix.py ./目录                     # 预览：哪些文件要转、是什么编码
  python3 csvfix.py ./目录 --apply             # 真转：输出到 ./目录/_utf8
  python3 csvfix.py ./目录 --apply --inplace   # 原地覆盖，自动留 .bak
  python3 csvfix.py a.csv b.csv --merge all.csv   # 合并成一份（表头只留一次）
  python3 csvfix.py ./目录 --merge all.csv --dedupe  # 合并同时去重
  python3 csvfix.py ./目录 --apply --undo csvfix_undo.json

不带 --apply 一律只预览，不动文件。
"""

import argparse
import csv
import hashlib
import io
import json
import os
import shutil
import sys
import time

ENCODINGS = ["utf-8-sig", "utf-8", "gb18030", "big5", "utf-16"]
BOM_UTF8 = b"\xef\xbb\xbf"


def sniff_encoding(raw):
    """返回 (编码名, 是否带BOM, 是否已是Excel友好)。gb18030 几乎能解所有字节，放最后兜底。"""
    if raw.startswith(BOM_UTF8):
        return "utf-8-sig", True, True
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return "utf-16", True, True
    for enc in ("utf-8", "gb18030", "big5"):
        try:
            raw.decode(enc)
            return enc, False, False
        except UnicodeDecodeError:
            continue
    return "gb18030", False, False


def sniff_delimiter(text):
    line = text.split("\n", 1)[0]
    counts = {d: line.count(d) for d in (",", "\t", ";", "|")}
    best = max(counts, key=counts.get)
    return best if counts[best] > 0 else ","


def read_rows(path):
    """读一个 CSV，返回 (行列表, 编码, 分隔符)。单元格全是字符串。"""
    with open(path, "rb") as f:
        raw = f.read()
    enc, has_bom, _ = sniff_encoding(raw)
    text = raw.decode(enc, errors="replace")
    delim = sniff_delimiter(text)
    rows = list(csv.reader(io.StringIO(text), delimiter=delim))
    return [r for r in rows if any((c or "").strip() for c in r)], enc, delim


def collect_files(targets):
    out = []
    for t in targets:
        if os.path.isdir(t):
            for root, _dirs, files in os.walk(t):
                if os.path.basename(root) == "_utf8":
                    continue
                for name in sorted(files):
                    if name.lower().endswith(".csv"):
                        out.append(os.path.join(root, name))
        elif os.path.isfile(t) and t.lower().endswith(".csv"):
            out.append(t)
    return out


def to_excel_friendly(src, dst, dedupe=False):
    raw = open(src, "rb").read()
    enc, has_bom, friendly = sniff_encoding(raw)
    text = raw.decode(enc, errors="replace")
    delim = sniff_delimiter(text)
    rows = list(csv.reader(io.StringIO(text), delimiter=delim))
    if dedupe:
        seen, kept = set(), []
        for r in rows:
            key = tuple((c or "").strip() for c in r)
            if key in seen:
                continue
            seen.add(key)
            kept.append(r)
        rows = kept
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=",", lineterminator="\n")
    w.writerows(rows)
    data = BOM_UTF8 + buf.getvalue().encode("utf-8")
    with open(dst, "wb") as f:
        f.write(data)
    return enc, friendly, len(rows), len(data)


def merge_files(files, out_path, dedupe=False, add_source=False):
    header, body, seen = None, [], set()
    per_file = []
    for f in files:
        rows, enc, _d = read_rows(f)
        if not rows:
            per_file.append((f, enc, 0))
            continue
        head, rest = rows[0], rows[1:]
        if header is None:
            header = head
        elif [c.strip() for c in head] != [c.strip() for c in header]:
            # 表头不一致就不当表头，整份当数据接在后面
            rest = rows
        got = 0
        for r in rest:
            key = tuple((c or "").strip() for c in r)
            if dedupe and key in seen:
                continue
            seen.add(key)
            if add_source:
                r = list(r) + [os.path.basename(f)]
            body.append(r)
            got += 1
        per_file.append((f, enc, got))
    if header is None:
        return per_file, 0
    if add_source:
        header = list(header) + ["来源文件"]
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=",", lineterminator="\n")
    w.writerow(header)
    w.writerows(body)
    with open(out_path, "wb") as fh:
        fh.write(BOM_UTF8 + buf.getvalue().encode("utf-8"))
    return per_file, len(body)


def do_undo(undo_path):
    data = json.load(open(undo_path, encoding="utf-8"))
    done = 0
    for it in data.get("actions", []):
        src, dst = it["src"], it["dst"]
        if it.get("mode") == "inplace":
            bak = src + ".bak"
            if os.path.exists(bak):
                shutil.move(bak, src)
                done += 1
        else:
            if os.path.exists(dst):
                os.remove(dst)
                done += 1
    print(f"已撤销 {done} 个文件")


def main():
    ap = argparse.ArgumentParser(description="CSV 乱码修复 / 合并 / 去重")
    ap.add_argument("targets", nargs="+", help="CSV 文件或目录，可传多个")
    ap.add_argument("--apply", action="store_true", help="真写盘（默认只预览）")
    ap.add_argument("--inplace", action="store_true", help="原地覆盖并留 .bak")
    ap.add_argument("--merge", metavar="OUT.csv", help="合并成一个文件")
    ap.add_argument("--dedupe", action="store_true", help="按整行去重")
    ap.add_argument("--source", action="store_true", help="合并时追加来源文件列")
    ap.add_argument("--undo", metavar="UNDO.json", help="按撤销记录回滚")
    args = ap.parse_args()

    if args.undo:
        do_undo(args.undo)
        return

    files = collect_files(args.targets)
    if not files:
        print("没找到 CSV 文件")
        sys.exit(1)

    if args.merge:
        per_file, total = merge_files(files, args.merge, args.dedupe, args.source)
        print(f"合并 {len(files)} 个文件 → {args.merge}（{total} 行数据，UTF-8-BOM，Excel 直接双击可看）")
        for f, enc, n in per_file:
            print(f"  {os.path.basename(f):<40} {enc:<10} {n} 行")
        return

    actions = []
    print(f"{'文件':<44}{'编码':<12}{'状态'}")
    for f in files:
        raw = open(f, "rb").read()
        enc, has_bom, friendly = sniff_encoding(raw)
        if friendly:
            print(f"{os.path.basename(f):<44}{enc:<12}已是 Excel 友好，跳过")
            continue
        if not args.apply:
            print(f"{os.path.basename(f):<44}{enc:<12}待转换")
            continue
        if args.inplace:
            dst = f
            shutil.copy2(f, f + ".bak")
            enc2, _fr, nrows, size = to_excel_friendly(f, dst, args.dedupe)
            actions.append({"src": f, "dst": f, "mode": "inplace"})
            print(f"{os.path.basename(f):<44}{enc2:<12}已原地转换（备份 .bak，{nrows} 行）")
        else:
            outdir = os.path.join(os.path.dirname(f), "_utf8")
            os.makedirs(outdir, exist_ok=True)
            dst = os.path.join(outdir, os.path.basename(f))
            enc2, _fr, nrows, size = to_excel_friendly(f, dst, args.dedupe)
            actions.append({"src": f, "dst": dst, "mode": "copy"})
            print(f"{os.path.basename(f):<44}{enc2:<12}已转换 → _utf8/{os.path.basename(f)}（{nrows} 行）")

    if args.apply and actions:
        undo = {"time": time.strftime("%Y-%m-%d %H:%M:%S"), "actions": actions}
        undo_path = "csvfix_undo.json"
        json.dump(undo, open(undo_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print(f"\n共处理 {len(actions)} 个文件，撤销记录写入 {undo_path}")
        print("改错了回滚：python3 csvfix.py . --undo csvfix_undo.json")
    elif not args.apply:
        print("\n以上只是预览。加 --apply 才会真写盘。")


if __name__ == "__main__":
    main()
