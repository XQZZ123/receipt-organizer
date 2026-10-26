# -*- coding: utf-8 -*-
"""按「日期-做什么-金额元.pdf」规范化重命名，并把发票归位到当月目录。

发票一律从当月目录（含 _raw 暂存子目录）里找，改名后平铺在 {month}/ 下，
_raw 里不再留副本 —— 同一张票只存在一份，避免下轮扫描把自己当重复。

用法：
    python organize.py --root <项目根目录> --month <YYYYMM> [--dry-run]
"""
import argparse
import glob
import json
import os
import re
import shutil
import sys

from scan_invoices import fmt_amount, init_buyer, load_excluded, parse_pdf

WHAT_BY_KIND = {"ride": "打车", "fuel": "加油", "meal": "吃饭", "hotel": "住宿", "rail": "高铁票"}
BRANDS = ["T3出行", "享道出行", "曹操出行", "如祺出行", "高德打车", "风韵出行", "滴滴出行"]


def what_word(rec):
    kind = rec.get("kind", "other")
    if kind == "rail":
        d = rec.get("detail", {})
        # 改签费票必须和行程票价票区分开，否则两行同名只差金额，财务看不出性质
        head = "高铁改签费" if d.get("is_change_fee") else "高铁票"
        if d.get("from") and d.get("to"):
            return f"{head}-{d['from']}-{d['to']}"
        return head
    if kind in WHAT_BY_KIND:
        return WHAT_BY_KIND[kind]
    # 其他：用票面项目名，如 *生产生活服务*技术服务 → 技术服务
    m = re.search(r"\*[^*]+\*(\S+)", rec.get("item", ""))
    return (m.group(1) if m else "其他").replace("服务费", "服务")


NAMED_RE = re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})-(.+)-([\d.]+)元$")


def already_organized(filename, rec):
    """文件名已是 日期-事项-金额元 结构、且日期金额与票面对得上 → 视为用户认可，不再改名。

    事项词是用户会手工调的（例：把「技术服务」改成「某SaaS费用」），
    自动改名会把手改覆盖回模板词，所以这里必须放过。
    """
    stem = os.path.splitext(filename)[0]
    m = NAMED_RE.match(stem)
    if not m:
        return False
    y, mo, d, _what, amt = m.groups()
    try:
        fdate = f"{int(y)}-{int(mo)}-{int(d)}"
    except ValueError:
        return False
    sdate = rec.get("service_date") or ""
    try:
        if float(amt) != float(rec.get("amount") or -1):
            return False
    except ValueError:
        return False
    return fdate == sdate


def brand_of(rec):
    name = os.path.basename(rec.get("file", ""))
    for b in BRANDS:
        if b in name:
            return b
    return ""


def target_name(rec):
    date = rec.get("service_date") or rec.get("invoice_date")
    amt = rec.get("amount")
    if not date or amt is None:
        return ""
    what = what_word(rec)
    return f"{date}-{what}-{fmt_amount(amt)}元.pdf"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--month", required=True)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    month_dir = os.path.join(os.path.abspath(args.root), args.month)
    raw_dir = os.path.join(month_dir, "_raw")
    excluded_dir = os.path.join(month_dir, "_excluded")
    init_buyer(os.path.abspath(args.root))
    excluded = load_excluded(os.path.abspath(args.root))

    planned, skipped, conflicts, parked = [], [], [], []
    seen = {}
    for p in sorted(glob.glob(os.path.join(month_dir, "**", "*.pdf"), recursive=True)):
        if excluded_dir in p:
            continue
        rec = parse_pdf(p)
        no = rec.get("invoice_no")
        if not no:
            skipped.append({"file": p, "reason": rec.get("skip_reason", "无发票号码")})
            continue
        # 用户登记为「已由他人报销/不入本表」→ 挪出平铺区，不参与出表也不拦后续扫描
        if no in excluded:
            dst = os.path.join(excluded_dir, os.path.basename(p))
            if not args.dry_run:
                os.makedirs(excluded_dir, exist_ok=True)
                if os.path.abspath(p) != os.path.abspath(dst):
                    shutil.move(p, dst)
            parked.append({"invoice_no": no, "src": p, "dst": dst,
                           "reason": excluded[no].get("reason", "")})
            print(f"  parked    {os.path.basename(p)[:46]:<46} → _excluded/"
                  f"（{excluded[no].get('reason', '已排除')}）")
            continue
        if no in seen:                       # 同号码已有文件：保留先出现的，源文件另存待删
            conflicts.append({"invoice_no": no, "keep": seen[no], "extra": p})
            continue
        seen[no] = p
        # 已在月份目录、且名字已符合规范（含用户手改的事项词）→ 原样保留
        if os.path.dirname(p) == month_dir and already_organized(os.path.basename(p), rec):
            planned.append({"src": p, "dst": p, "invoice_no": no, "brand": brand_of(rec),
                            "action": "unchanged"})
            continue
        new_name = target_name(rec)
        if not new_name:
            skipped.append({"file": p, "reason": "缺日期或金额，无法命名"})
            continue
        planned.append({"src": p, "dst": os.path.join(month_dir, new_name), "invoice_no": no,
                        "brand": brand_of(rec)})

    # 目标名已被占用且不是同一个文件 → 加序号
    taken = {os.path.basename(p["dst"]) for p in planned}
    for item in planned:
        dst = item["dst"]
        if os.path.exists(dst) and os.path.abspath(dst) != os.path.abspath(item["src"]):
            base, ext = os.path.splitext(dst)
            i = 2
            while os.path.exists(f"{base}-{i}{ext}") or f"{os.path.basename(base)}-{i}{ext}" in taken:
                i += 1
            item["dst"] = f"{base}-{i}{ext}"

    for item in planned:
        src, dst = item["src"], item["dst"]
        if item.get("action") == "unchanged" or os.path.abspath(src) == os.path.abspath(dst):
            item["action"] = "unchanged"
            continue
        item["action"] = "rename" if os.path.dirname(src) == month_dir else "move"
        if not args.dry_run:
            if os.path.exists(dst):
                os.remove(dst)               # 同号码目标已存在且内容重复，覆盖即可
            shutil.move(src, dst)
        print(f"  {item['action']:<9} {os.path.basename(src)[:46]:<46} → {os.path.basename(dst)}")

    # 清理空的暂存目录
    if not args.dry_run and os.path.isdir(raw_dir) and not os.listdir(raw_dir):
        os.rmdir(raw_dir)

    print(json.dumps({
        "month": args.month,
        "renamed": len([p for p in planned if p["action"] != "unchanged"]),
        "unchanged": len([p for p in planned if p["action"] == "unchanged"]),
        "parked_excluded": parked,
        "skipped": skipped,
        "duplicate_files": conflicts,
        "dry_run": args.dry_run,
    }, ensure_ascii=False, indent=2))

    if conflicts:
        print("!! 同一发票号码存在多份文件，已保留第一份，其余未动：", file=sys.stderr)
        for c in conflicts:
            print(f"   {c['invoice_no']} 多余文件: {c['extra']}", file=sys.stderr)


if __name__ == "__main__":
    main()
