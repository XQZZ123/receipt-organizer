# -*- coding: utf-8 -*-
"""把「QQ 邮箱附件下载」的本地结果文件解码成真实 PDF。

宿主会把每次工具调用的原始输出存在
    <project>/<session>/tool-results/call_*.txt        （原文）
    <project>/<session>/tool-results/call_*-render.txt （渲染副本，内容相同）
内容形如
    {"data":{"content":"<base64>","filename":"xxx.pdf|xxx.zip","content_type":"..."}}

两种附件都要处理：
  - 直接是 PDF：校验 %PDF 头后落盘；
  - 是 ZIP（12306 的电子发票就是 zip 里装 pdf/ofd/xml）：解压后只取 PDF。
全程不经过模型上下文，靠 content 的 SHA-1 去重（原文与 render 副本会重复）。

用法：
    python fetch_attachments.py --out <当月目录>/_raw [--since-minutes 60] [--keep-ext .pdf]
"""
import argparse
import base64
import glob
import hashlib
import io
import json
import os
import sys
import time
import zipfile

SKIP_SUFFIX = "-render.txt"


def iter_payloads(paths):
    """产出 (filename, blob)；同一 base64 只产出一次。"""
    seen = set()
    for p in paths:
        try:
            with open(p, encoding="utf-8", errors="replace") as fh:
                text = fh.read().strip()
        except OSError:
            continue
        if not text.startswith("{"):
            continue
        try:
            obj = json.loads(text)
        except json.JSONDecodeError:
            continue
        data = obj.get("data") if isinstance(obj, dict) else None
        if not isinstance(data, dict):
            continue
        content, name = data.get("content"), data.get("filename")
        if not content or not name:
            continue
        digest = hashlib.sha1(content.encode()).hexdigest()
        if digest in seen:
            continue
        seen.add(digest)
        try:
            blob = base64.b64decode(content, validate=False)
        except Exception:
            continue
        yield name, blob


def extract_pdfs(name, blob):
    """ZIP → 内部的 PDF 列表；单个 PDF → 自己。返回 [(filename, bytes)]。"""
    if zipfile.is_zipfile(io.BytesIO(blob)):
        out = []
        with zipfile.ZipFile(io.BytesIO(blob)) as zf:
            for info in zf.infolist():
                if info.is_dir():
                    continue
                inner = zf.read(info)
                if inner[:4] == b"%PDF":
                    out.append((os.path.basename(info.filename), inner))
        return out
    if blob[:4] == b"%PDF":
        return [(name, blob)]
    return []


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", nargs="*", default=None,
                    help="tool-results 目录；省略则自动探测所有项目下的目录")
    ap.add_argument("--out", required=True, help="落盘目录")
    ap.add_argument("--since-minutes", type=float, default=0,
                    help="只看最近 N 分钟内生成的结果文件；0 = 不限")
    args = ap.parse_args()

    os.makedirs(args.out, exist_ok=True)
    cutoff = time.time() - args.since_minutes * 60 if args.since_minutes else 0

    src_dirs = args.src
    if not src_dirs:
        base = os.path.expanduser(os.path.join("~", ".workbuddy", "projects"))
        src_dirs = glob.glob(os.path.join(base, "*", "*", "tool-results"))
    if not src_dirs:
        print("找不到 tool-results 目录，请用 --src 指定", file=sys.stderr)
        sys.exit(1)

    files = []
    for d in src_dirs:
        for p in glob.glob(os.path.join(d, "*.txt")):
            if p.endswith(SKIP_SUFFIX):
                continue            # render 副本，与原文重复
            if cutoff and os.path.getmtime(p) < cutoff:
                continue
            files.append(p)
    files.sort(key=os.path.getmtime)

    saved, skipped = [], []
    for att_name, blob in iter_payloads(files):
        pdfs = extract_pdfs(att_name, blob)
        if not pdfs:
            skipped.append(f"{att_name}（非 PDF/ZIP 或无内嵌 PDF）")
            continue
        for inner_name, inner in pdfs:
            dst = os.path.join(args.out, inner_name)
            if os.path.exists(dst) and os.path.getsize(dst) == len(inner):
                skipped.append(f"{inner_name}（已存在）")
                continue
            with open(dst, "wb") as fh:
                fh.write(inner)
            saved.append({"file": dst, "bytes": len(inner), "from": att_name})

    print(json.dumps({"saved": saved, "skipped": skipped}, ensure_ascii=False, indent=2))
    if not saved:
        print("没有新落盘的 PDF；确认 DownloadAttachment 是否真的跑过、时间窗是否够。",
              file=sys.stderr)


if __name__ == "__main__":
    main()
