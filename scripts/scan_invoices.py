# -*- coding: utf-8 -*-
"""扫描发票 PDF，抽字段 + 跨月查重，输出 JSON 与可读清单。

只在 pdf 文本层做，不做 OCR。已处理三个真实踩过的坑：
  1. 康熙部首异体字：票面 PDF 常把「长春/金拱门/生产」写成 ⻓/⾦/⽣，必须 NFKC 归一化；
  2. 字段顺序因开票厂商而异（高德系 4 家各不同），一律按特征匹配，不按位置；
  3. 价税合计优先用中文大写金额解析，¥ 面值取 max 兜底（合计=金额+税额，必为最大项）。

用法：
    python scan_invoices.py --root <项目根目录> --month <YYYYMM> [--json out.json]
"""
import argparse
import glob
import json
import os
import re
import sys
import unicodedata
from datetime import date

import pymupdf

# 购买方抬头由 .invoice-ledger/config.json 的 buyer 提供，不写死在代码里（换公司/换人可复用）。
# 为空时跳过抬头相关校验与过滤，不会误报。
BUYER = ""
MONTH_DIR_RE = re.compile(r"^\d{6}$")
INV_NO_RE = re.compile(r"(?<!\d)(\d{20})(?!\d)")
CN_DATE_RE = re.compile(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日")
ISO_DATE_RE = re.compile(r"(20\d{2})-(\d{1,2})-(\d{1,2})")
MONEY_RE = re.compile(r"[¥￥]\s*([\d,]+\.\d{2})")
ITEM_RE = re.compile(r"\*([^*\n]{2,20})\*([^*\n]{2,30})")
CO_NAME_RE = re.compile(
    r"[\u4e00-\u9fa5A-Za-z0-9（）()·]{2,40}"
    r"(?:股份有限公司|有限责任公司|有限公司|分公司|大学|研究院|事务所)"
)
# 票面标签词会把公司名粘成「名称XX公司…」，先替换成分隔符再匹配
LABEL_RE = re.compile(
    r"(?:购买方名称|销售方名称|购买方|销售方|名称|统一社会信用代码[/／]纳税人识别号|纳税人识别号)[:：]?"
)
TRAIN_STATION_RE = re.compile(r"([\u4e00-\u9fa5]{2,8}站)")
TRAIN_NO_RE = re.compile(r"(?<![0-9A-Za-z])([GDCKZT]\d{1,4})(?![0-9])")

CN_D = {"零": 0, "〇": 0, "壹": 1, "贰": 2, "叁": 3, "肆": 4, "伍": 5,
        "陆": 6, "柒": 7, "捌": 8, "玖": 9, "两": 2}
CN_U = {"拾": 10, "佰": 100, "仟": 1000}
CN_AMT_RE = re.compile(r"[零〇壹贰叁肆伍陆柒捌玖拾佰仟万]{1,}[元圆][零〇壹贰叁肆伍陆柒捌玖角分整]{0,}")

# 票面线索 → (费用科目, 事项词)
CATEGORY_RULES = [
    ("铁路电子客票", "交通费", "rail"),
    ("成品油", "交通费", "fuel"),
    ("汽油", "交通费", "fuel"),
    ("柴油", "交通费", "fuel"),
    ("住宿服务", "住宿费", "hotel"),
    ("餐饮服务", "餐饮费", "meal"),
    ("旅客运输服务", "交通费", "ride"),
    ("客运服务费", "交通费", "ride"),
    ("交通运输服务", "交通费", "ride"),
    ("技术服务", "其他费用", "other"),
]


def norm(text):
    """NFKC 归一化：康熙部首/全角字符 → 常规汉字；统一换行。"""
    text = unicodedata.normalize("NFKC", text)
    return text.replace("\r\n", "\n").replace("\r", "\n")


def parse_cn_amount(s):
    """中文大写金额 → float。支持「贰拾肆元捌角陆分」「伍佰叁拾柒圆陆角伍分」「贰拾叁圆伍角整」。"""
    m = re.match(r"(.+?)[元圆](.*)$", s)
    if not m:
        return None
    int_part, frac = m.group(1), m.group(2)
    total = section = num = 0
    for ch in int_part:
        if ch == "万":
            total += (section + num) * 10000
            section = num = 0
        elif ch in CN_U:
            num = (num or 1) * CN_U[ch]
            section += num
            num = 0
        elif ch in CN_D:
            num = CN_D[ch]
    yuan = total + section + num

    jiao = fen = 0
    mj = re.search(r"([零〇壹贰叁肆伍陆柒捌玖])角", frac)
    if mj:
        jiao = CN_D[mj.group(1)]
    mf = re.search(r"([零〇壹贰叁肆伍陆柒捌玖])分", frac)
    if mf:
        fen = CN_D[mf.group(1)]
    return round(yuan + jiao / 10 + fen / 100, 2)


def fmt_amount(v):
    """437.00→437，28.80→28.8，27.99→27.99（与既有文件名口径一致）。"""
    return f"{v:.2f}".rstrip("0").rstrip(".")


def fmt_date(y, m, d):
    """2026-9-18 不补零，与既有文件名口径一致。"""
    return f"{y}-{m}-{d}"


def pick_seller(text):
    cleaned = LABEL_RE.sub("|", text)
    names = CO_NAME_RE.findall(cleaned)
    out = []
    for n in names:
        n = n.strip("| :：")
        if not n or n in out:
            continue
        if BUYER and (n == BUYER or BUYER in n):
            continue
        out.append(n)
    # 包容关系只留最长的（「某某石油销售股份有限公司」与「…某某石油分公司」实为同一家）
    out = [n for n in out if not any(n != o and n in o for o in out)]
    if not out:
        return "", []
    out.sort(key=len, reverse=True)
    return out[0], out[1:]


def parse_pdf(path):
    doc = pymupdf.open(path)
    raw = "\n".join(page.get_text() for page in doc)
    doc.close()
    text = norm(raw)
    flat = re.sub(r"[ \t]+", " ", text)
    warnings = []

    # 行程单 / 非发票
    if "行程单" in os.path.basename(path) and not INV_NO_RE.search(flat):
        return {"file": path, "kind": "itinerary", "skip_reason": "行程单，不是发票"}

    inv_no = ""
    m = re.search(r"发票号码[:：]?\s*(\d{20})", flat)
    inv_no = m.group(1) if m else (INV_NO_RE.search(flat).group(1) if INV_NO_RE.search(flat) else "")
    if not inv_no:
        return {"file": path, "kind": "unknown", "skip_reason": "未识别出发票号码，可能是行程单/图片型 PDF"}

    # 开票日期
    inv_date = ""
    m = CN_DATE_RE.search(flat)
    if m:
        inv_date = fmt_date(int(m.group(1)), int(m.group(2)), int(m.group(3)))

    # 金额：大写优先，¥ 取 max 兜底
    amt_cn = None
    m = CN_AMT_RE.search(flat)
    if m:
        amt_cn = parse_cn_amount(m.group(0))
    moneys = [float(x.replace(",", "")) for x in MONEY_RE.findall(flat)]
    amt_max = max(moneys) if moneys else None
    amount = amt_cn if amt_cn is not None else amt_max
    if amt_cn is not None and amt_max is not None and abs(amt_cn - amt_max) > 0.005:
        warnings.append(f"大写金额 {amt_cn} 与票面最大¥值 {amt_max} 不一致，已采用大写")

    seller, other_sellers = pick_seller(flat)
    if not BUYER:
        # 票面上只有「购买方」「销售方」两个公司名，没有抬头就无法判断哪个是卖方。
        # 此时宁可留空让用户去配置，也绝不猜 —— 猜错会把自家公司名写成开票单位，静默污染整张表。
        if seller or other_sellers:
            warnings.append("未配置 buyer，无法区分购买方/销售方，开票单位已留空")
            seller = ""
        other_sellers = []
    elif other_sellers:
        warnings.append("销售方候选不唯一: " + " / ".join([seller] + other_sellers))
    if BUYER and BUYER not in flat:
        warnings.append(f"票面未出现抬头「{BUYER}」，确认是否应报销")

    items = ITEM_RE.findall(flat)
    item_text = "；".join(f"*{a}*{b}" for a, b in items)

    kind, category = "other", "待定"
    for kw, cat, kd in CATEGORY_RULES:
        if kw in flat:
            category, kind = cat, kd
            break

    # 出行日期：ISO 日期里排除开票日期；多行程票取最早一天作为报销挂账日
    ymd = {(int(y), int(mo), int(d)) for y, mo, d in ISO_DATE_RE.findall(flat)}
    inv_ymd = tuple(int(x) for x in re.findall(r"\d+", inv_date)) if inv_date else None
    if inv_ymd:
        ymd.discard(inv_ymd)
    trips = [fmt_date(*o) for o in sorted(ymd)]

    detail = {}
    if kind == "rail":
        nos = TRAIN_NO_RE.findall(flat)
        st = TRAIN_STATION_RE.findall(flat)
        seat = re.search(r"(\d{1,2}车\s*\d{1,3}[A-Z]?号)", flat)
        cls = re.search(r"(商务座|一等座|二等座|无座|硬座|软卧|硬卧)", flat)
        m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日\s*\d{2}:\d{2}开", flat)
        if m:
            trips = [fmt_date(int(m.group(1)), int(m.group(2)), int(m.group(3)))]
        detail = {
            "train": nos[0] if nos else "",
            "from": st[0].rstrip("站") if len(st) > 0 else "",
            "to": st[1].rstrip("站") if len(st) > 1 else "",
            "seat": seat.group(1).replace(" ", "") if seat else "",
            "class": cls.group(1) if cls else "",
            # 改签费票：票面「改签费:」而非「票价:」，金额只是手续费，不是行程票价
            "is_change_fee": "改签费:" in flat or "改签费：" in flat,
        }

    service_date = trips[0] if trips else inv_date
    # 铁路电子客票票面没有销售方公司名，只有「中国铁路祝您旅途愉快」，
    # 沿用 9 月定稿写法：中国铁路（出发站）
    if kind == "rail" and not seller:
        dep = detail.get("from", "")
        seller = f"中国铁路（{dep}站）" if dep else "中国铁路"
    return {
        "file": path,
        "kind": kind,
        "category": category,
        "invoice_no": inv_no,
        "invoice_date": inv_date,
        "service_date": service_date,
        "trip_dates": trips,
        "amount": amount,
        "seller": seller,
        "item": item_text,
        "detail": detail,
        "warnings": warnings,
    }


def scan(dirpath):
    out = []
    for p in sorted(glob.glob(os.path.join(dirpath, "**", "*.pdf"), recursive=True)):
        try:
            out.append(parse_pdf(p))
        except Exception as exc:  # 单张坏文件不拖垮整批
            out.append({"file": p, "kind": "error", "skip_reason": f"解析失败: {exc}"})
    return out


LEDGER_DIR = ".invoice-ledger"
CONFIG_NAME = "config.json"

DEFAULT_CONFIG = {
    "buyer": "",
    "cats": ["交通费", "住宿费", "餐饮费", "其他费用"],
}


def load_config(root, create=False):
    """读 .invoice-ledger/config.json；create=True 时缺失则写一份模板出来。

    buyer: 本公司抬头全称，用于「过滤掉自己公司名」和「校验票面抬头」。
           留空则这两项检查都跳过（不会误报，但也不会提醒抬头不对）。
    cats : 费用科目分区顺序，需与 CATEGORY_RULES 里的中文名一致。
    """
    d = os.path.join(root, LEDGER_DIR)
    p = os.path.join(d, CONFIG_NAME)
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(p):
        try:
            with open(p, encoding="utf-8") as fh:
                cfg.update({k: v for k, v in json.load(fh).items() if not k.startswith("_")})
        except Exception:
            pass
    elif create:
        os.makedirs(d, exist_ok=True)
        with open(p, "w", encoding="utf-8") as fh:
            json.dump({"_说明": "buyer 填本公司发票抬头全称；cats 是费用科目分区顺序，改这里也要改 CATEGORY_RULES。",
                       **cfg}, fh, ensure_ascii=False, indent=2)
    return cfg


def init_buyer(root):
    """把 config 里的抬头装进模块级 BUYER。各入口脚本 main() 里调一次即可。"""
    global BUYER
    BUYER = (load_config(root).get("buyer") or "").strip()
    return BUYER


def load_excluded(root):
    """读「已由他人报销/不入本表」登记。号码为键，值为原因等元信息。"""
    p = os.path.join(root, LEDGER_DIR, "excluded.json")
    if not os.path.exists(p):
        return {}
    try:
        with open(p, encoding="utf-8") as fh:
            data = json.load(fh)
    except Exception:
        return {}
    return {k: v for k, v in data.items()
            if re.fullmatch(r"\d{20}", k) and isinstance(v, dict)}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True, help="项目根目录")
    ap.add_argument("--month", required=True, help="当月目录名，形如 YYYYMM")
    ap.add_argument("--json", default="", help="明细 JSON 输出路径")
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    buyer = init_buyer(root)
    if not buyer:
        print(f"提示：未配置公司抬头（{LEDGER_DIR}/{CONFIG_NAME} 的 buyer），"
              f"已跳过抬头校验与自身公司名过滤。")
    target_dir = os.path.join(root, args.month)
    targets = scan(target_dir)

    # 跨月查重：历史月份目录只取发票号码
    history = {}
    for d in sorted(glob.glob(os.path.join(root, "*"))):
        name = os.path.basename(d)
        if not os.path.isdir(d) or not MONTH_DIR_RE.match(name) or name == args.month:
            continue
        for rec in scan(d):
            if rec.get("invoice_no"):
                history.setdefault(rec["invoice_no"], []).append(name)

    # 已记录的发票号码，分两类：
    #   历史月份/往年表 → 真·重复报销风险，必须问用户
    #   当月表         → 只是「已录入」，不是重复，本轮增量整理要跳过而不是告警
    current_sheet = {}
    xlsx_files = [p for p in glob.glob(os.path.join(root, "**", "*.xlsx"), recursive=True)
                  if not os.path.basename(p).startswith("~$")]
    for xlsx in xlsx_files:
        in_current_month = os.path.dirname(os.path.abspath(xlsx)) == os.path.join(root, args.month)
        try:
            import openpyxl
            wb = openpyxl.load_workbook(xlsx, data_only=True)
        except Exception:
            continue
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                for cell in row:
                    s = str(cell or "")
                    if re.fullmatch(r"\d{20}", s):
                        if in_current_month:
                            current_sheet.setdefault(s, []).append(os.path.basename(xlsx))
                        else:
                            history.setdefault(s, []).append(os.path.basename(xlsx))

    # 用户已登记「由他人报销/不入本表」的发票：号码进黑名单，不算新票、不告警
    excluded = load_excluded(root)

    valid = [r for r in targets if r.get("invoice_no")]
    seen = {}
    dups, hists, recorded, excluded_hits = [], [], [], []
    for r in valid:
        no = r["invoice_no"]
        if no in seen:
            dups.append({"invoice_no": no, "files": [seen[no], r["file"]]})
        else:
            seen[no] = r["file"]
        if no in excluded:
            excluded_hits.append({"invoice_no": no, "file": r["file"],
                                  "reason": excluded[no].get("reason", "")})
            continue
        if no in history:
            hists.append({"invoice_no": no, "file": r["file"], "seen_in": sorted(set(history[no]))})
        if no in current_sheet:
            recorded.append({"invoice_no": no, "file": r["file"],
                             "sheet": sorted(set(current_sheet[no]))})

    new_invoices = [r for r in valid if r["invoice_no"] not in current_sheet
                    and r["invoice_no"] not in history
                    and r["invoice_no"] not in excluded]

    report = {
        "month": args.month,
        "target_count": len(valid),
        "invoices": valid,
        "new_invoices": new_invoices,
        "skipped": [r for r in targets if not r.get("invoice_no")],
        "duplicates_in_target": dups,
        "already_reported": hists,
        "already_in_current_sheet": recorded,
        "excluded_by_user": excluded_hits,
        "history_invoice_count": len(history),
        "current_sheet_invoice_count": len(current_sheet),
    }

    if args.json:
        with open(args.json, "w", encoding="utf-8") as fh:
            json.dump(report, fh, ensure_ascii=False, indent=2)

    print(f"当月 {args.month}：识别 {len(valid)} 张，跳过 {len(report['skipped'])} 个文件")
    for r in valid:
        amt = r["amount"] if r["amount"] is not None else 0
        flag = ""
        if r["invoice_no"] in excluded:
            flag = f"  [勿录入：{excluded[r['invoice_no']].get('reason', '已排除')}]"
        elif r["invoice_no"] in current_sheet:
            flag = "  [已录入本月表]"
        print(f"  {r['service_date']:>10} | {r['category']:<4} | {amt:>8.2f} | {r['invoice_no']} | "
              f"{r['seller'][:22]:<22} | {os.path.basename(r['file'])}{flag}")
    print(f"\n本轮新增待录入：{len(new_invoices)} 张")
    if excluded_hits:
        print("\n已登记排除（不计入本表）：")
        for h in excluded_hits:
            print(f"  {h['invoice_no']} — {h['reason']}（{os.path.basename(h['file'])}）")
    if report["skipped"]:
        print("\n跳过：")
        for r in report["skipped"]:
            print(f"  {os.path.basename(r['file'])} — {r.get('skip_reason')}")
    if dups:
        print("\n!! 当月内部重复（同发票号码多文件）：")
        for d in dups:
            if d["invoice_no"] in excluded:
                continue
            print(f"  {d['invoice_no']}: " + " , ".join(os.path.basename(f) for f in d["files"]))
    if hists:
        print("\n!! 历史月份已报过（重复报销风险，务必确认）：")
        for h in hists:
            print(f"  {h['invoice_no']} 已在 {', '.join(h['seen_in'])}")
    if not dups and not hists:
        print(f"\n查重通过：与历史 {len(history)} 张已录发票无重叠。")
    warns = [(r["invoice_no"], w) for r in valid for w in r.get("warnings", [])]
    if warns:
        print("\n提示：")
        for no, w in warns:
            print(f"  {no}: {w}")


if __name__ == "__main__":
    main()
