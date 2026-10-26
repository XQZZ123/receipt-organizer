# -*- coding: utf-8 -*-
"""交付前校验：错误值 / 缓存值 / 公式保留 / 张数对账 / 号码去重 / 小计合计相符。

任何一项 FAIL 都不要交付。用法：
    python check_delivery.py --xlsx <月份目录>/<当月表>.xlsx --month-dir <月份目录>
"""
import argparse
import glob
import os
import re
import sys

import openpyxl
from openpyxl.utils import get_column_letter

ERR = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NULL!", "#NUM!")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--month-dir", required=True)
    args = ap.parse_args()

    results = []
    wv = openpyxl.load_workbook(args.xlsx, data_only=True)
    wf = openpyxl.load_workbook(args.xlsx)
    ws, wsf = wv.worksheets[0], wf.worksheets[0]

    # 错误值
    bad = [(ws.title, c.coordinate, c.value) for row in ws.iter_rows() for c in row
           if isinstance(c.value, str) and any(e in c.value for e in ERR)]
    results.append(("无错误值", not bad, bad or "ok"))

    # 缓存值：小计/合计行的金额必须读得出来（否则预览器一片空白）
    amounts, formulas, nos = [], [], []
    for r in range(1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        e = ws.cell(r, 5).value
        f = ws.cell(r, 6).value
        if isinstance(a, str) and ("小计" in a or "合" in a):
            amounts.append((r, a, e))
            fv = wsf.cell(r, 1).value
            if isinstance(fv, str) and fv.startswith("="):
                formulas.append(r)
        if re.fullmatch(r"\d{20}", str(f or "")):
            nos.append(str(f))
    missing = [x for x in amounts if not isinstance(x[2], (int, float))]
    results.append(("小计/合计缓存值齐全", not missing, missing or "ok"))
    results.append(("小计/合计仍保留公式", len(formulas) >= len(amounts), f"{len(formulas)}/{len(amounts)} 行"))

    # 张数对账：月份目录里平铺的 PDF 数 vs 表里发票行数
    # 只算平铺，_raw/_excluded 等子目录不算（被排除的票按规则移出平铺区）
    pdfs = [p for p in glob.glob(os.path.join(args.month_dir, "*.pdf"))]
    results.append(("张数 = 目录 PDF 数", len(pdfs) == len(nos), f"PDF {len(pdfs)} 张 / 表中 {len(nos)} 行"))

    # 被用户排除的票不得出现在表里
    ledger = os.path.join(os.path.dirname(os.path.abspath(args.month_dir)), ".invoice-ledger", "excluded.json")
    if os.path.exists(ledger):
        import json
        with open(ledger, encoding="utf-8") as fh:
            excluded = {k for k in json.load(fh) if re.fullmatch(r"\d{20}", k)}
        leaked = sorted(excluded & set(nos))
        results.append(("已排除发票未混入表中", not leaked, leaked or "ok"))

    # 发票号码唯一
    dup = sorted({n for n in nos if nos.count(n) > 1})
    results.append(("发票号码唯一", not dup, dup or f"{len(set(nos))} 个唯一号码"))

    # 小计相加 = 合计
    subs = [x[2] for x in amounts if "小计" in str(x[1])]
    total = [x[2] for x in amounts if "合" in str(x[1])]
    if subs and total and all(isinstance(v, (int, float)) for v in subs + total):
        s = round(sum(subs), 2)
        ok = abs(s - round(total[0], 2)) < 0.005
        results.append(("各块小计相加 = 合计", ok, f"小计和 {s} / 合计 {total[0]}"))
    else:
        results.append(("各块小计相加 = 合计", False, "读不到小计或合计"))

    print(f"表结构：{ws.max_row} 行 × {ws.max_column} 列")
    print("明细：")
    for r in range(3, ws.max_row + 1):
        cells = [ws.cell(r, c).value for c in range(1, 9)]
        if all(v is None for v in cells):
            continue
        print(" ", " | ".join("" if v is None else str(v)[:30] for v in cells))

    print()
    for name, ok, detail in results:
        print(f"  [{'PASS' if ok else 'FAIL'}] {name}: {detail}")
    if not all(ok for _, ok, _ in results):
        sys.exit(1)


if __name__ == "__main__":
    main()
