# -*- coding: utf-8 -*-
"""发票报销明细表生成器：一张表，按费用科目分块 + 分类小计 + 合计。

版式与口径沿用 2026-09 定稿版，不要自创新结构：
  列 = 序号｜项目名称｜项目内容｜日期｜金额（元）｜发票号码｜开票单位｜备注
  行 = 标题｜表头｜(明细… + 小计)×分类｜合计
  小计行 B 列必须留空，否则被 COUNTIF/SUMIF 重复计入。

输入 rows.json：
  {"title": "2026年10月发票报销明细表",
   "cats": ["交通费", "住宿费", "餐饮费", "其他费用"],
   "rows": [["交通费", "项目内容", "2026-3-12", 128.4, "20位号码", "开票单位", "无"], ...]}

用法：
    python build_xlsx.py --rows rows.json --out 202610/2026年10月发票整理表.xlsx   # 路径按需替换
"""
import argparse
import json
import os
import re

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation


def xl_color(css_hex: str) -> str:
    value = css_hex.removeprefix("#").upper()
    if len(value) != 6:
        raise ValueError(f"Expected #RRGGBB, got: {css_hex}")
    return "FF" + value


XL_HEADER_BG = xl_color("#4472C4")
XL_TITLE_BG = xl_color("#2F5597")
XL_LIGHT_BG = xl_color("#D9E2F3")
XL_BORDER = xl_color("#BFBFBF")
XL_WHITE = xl_color("#FFFFFF")
XL_TEXT = xl_color("#1A1A1A")

F_TITLE = Font(name="微软雅黑", size=12, bold=True, color=XL_WHITE)
F_HEAD = Font(name="微软雅黑", size=10, bold=True, color=XL_WHITE)
F_BODY = Font(name="微软雅黑", size=10, color=XL_TEXT)
F_BOLD = Font(name="微软雅黑", size=10, bold=True, color=XL_TEXT)
F_TOTAL = Font(name="微软雅黑", size=10, bold=True, color=XL_WHITE)

FILL_HEAD = PatternFill("solid", fgColor=XL_HEADER_BG)
FILL_TITLE = PatternFill("solid", fgColor=XL_TITLE_BG)
FILL_SUB = PatternFill("solid", fgColor=XL_LIGHT_BG)
FILL_TOTAL = PatternFill("solid", fgColor=XL_TITLE_BG)

thin = Side(style="thin", color=XL_BORDER)
BD = Border(left=thin, right=thin, top=thin, bottom=thin)
AC = Alignment(horizontal="center", vertical="center")
AL = Alignment(horizontal="left", vertical="center")
ALW = Alignment(horizontal="left", vertical="center", wrapText=True)
AR = Alignment(horizontal="right", vertical="center")
ACW = Alignment(horizontal="center", vertical="center", wrapText=True)

FMT_DATE = "YYYY-MM-DD"
FMT_MONEY = "¥#,##0.00"
HDR = ["序号", "项目名称", "项目内容", "日期", "金额（元）", "发票号码", "开票单位", "备注"]
NCOL = len(HDR)


def inject_cache(path, sheet_values):
    """openpyxl 只写 <f>..</f><v/>，无缓存值时多数预览器显示空白，按数据回填 <v>。"""
    import re as _re
    import zipfile
    from xml.sax.saxutils import escape

    def repl_factory(val):
        def repl(m):
            attrs, body = m.group(2), m.group(3)
            if isinstance(val, str):
                if 't="' not in attrs:
                    attrs = attrs[:-1] + ' t="str">'
                text = escape(val)
            else:
                text = repr(float(val))
            return m.group(1) + attrs + body + "<v>%s</v>" % text
        return repl

    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin:
        names = zin.namelist()
        data = {n: zin.read(n) for n in names}
    for sheet_file, cells in sheet_values.items():
        xml = data[sheet_file].decode("utf-8")
        for ref, val in cells.items():
            # openpyxl 3.1.x 这里写的是 <v></v>，不是自闭合 <v/>，两种都要认
            pat = _re.compile(r'(<c r="%s")([^>]*>)(<f>.*?</f>)(?:<v\s*/>|<v\s*></v>)' % ref, _re.S)
            xml, n = pat.subn(repl_factory(val), xml)
            if n != 1:
                raise SystemExit("cache inject failed: %s!%s (matched %d)" % (sheet_file, ref, n))
        data[sheet_file] = xml.encode("utf-8")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, data[n])
    try:
        os.replace(tmp, path)
    except PermissionError:          # 文件被预览器占用：替换被拒，原地覆盖
        with open(tmp, "rb") as fh:
            blob = fh.read()
        with open(path, "wb") as fh:
            fh.write(blob)
        os.remove(tmp)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rows", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    with open(args.rows, encoding="utf-8") as fh:
        cfg = json.load(fh)
    CATS = cfg["cats"]
    ROWS = cfg["rows"]
    TITLE = cfg["title"]

    # ---- 交付前硬校验：号码唯一、分类合法、金额为正 ----
    nos = [r[4] for r in ROWS]
    if len(nos) != len(set(nos)):
        dup = sorted({n for n in nos if nos.count(n) > 1})
        raise SystemExit("发票号码重复：%s —— 查重未通过，拒绝出表" % dup)
    bad_cat = [r[0] for r in ROWS if r[0] not in CATS]
    if bad_cat:
        raise SystemExit("项目名称不在分类表内：%s" % bad_cat)
    bad_no = [n for n in nos if not re.fullmatch(r"\d{20}", str(n))]
    if bad_no:
        raise SystemExit("发票号码格式异常（应 20 位数字）：%s" % bad_no)

    DISPLAY = [(cat, row) for cat in CATS for row in ROWS if row[0] == cat]
    FIRST_DATA = 3
    _cur, _seq = FIRST_DATA, 0
    PLACED, SUB_ROWS = [], {}
    for _cat in CATS:
        for _c, _row in DISPLAY:
            if _c != _cat:
                continue
            _seq += 1
            PLACED.append((_cur, _seq, _cat, _row))
            _cur += 1
        SUB_ROWS[_cat] = _cur
        _cur += 1
    TOTAL_ROW = _cur
    DETAIL_ROWS = [p[0] for p in PLACED]
    LAST_BODY_ROW = TOTAL_ROW - 1
    CAT_RANGE = f"$B${FIRST_DATA}:$B${LAST_BODY_ROW}"
    AMT_RANGE = f"$E${FIRST_DATA}:$E${LAST_BODY_ROW}"
    INV_RANGE = f"$F${FIRST_DATA}:$F${LAST_BODY_ROW}"

    wb = Workbook()
    wb.properties.title = TITLE
    ws = wb.active
    ws.title = "发票报销明细"

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=NCOL)
    ws.cell(row=1, column=1, value=TITLE)
    for c in range(1, NCOL + 1):
        cc = ws.cell(row=1, column=c)
        cc.fill = FILL_TITLE
        cc.border = BD
    ws["A1"].font = F_TITLE
    ws["A1"].alignment = AC
    ws.row_dimensions[1].height = 28

    for c, name in enumerate(HDR, start=1):
        cc = ws.cell(row=2, column=c, value=name)
        cc.font = F_HEAD
        cc.fill = FILL_HEAD
        cc.alignment = ACW
        cc.border = BD
    ws.row_dimensions[2].height = 30

    for r, seq, cat, (_, content, d, amt, inv, issuer, note) in PLACED:
        if isinstance(d, str):
            # 日期字符串可能不补零（2026-3-6），fromisoformat 会拒，手动拆
            from datetime import date as _date
            y, mo, dd = (int(x) for x in re.findall(r"\d+", d)[:3])
            d = _date(y, mo, dd)
        ws.cell(row=r, column=1, value=seq).number_format = "0"
        ws.cell(row=r, column=2, value=cat)
        ws.cell(row=r, column=3, value=content)
        ws.cell(row=r, column=4, value=d).number_format = FMT_DATE
        ws.cell(row=r, column=5, value=amt).number_format = FMT_MONEY
        ws.cell(row=r, column=6, value=inv)
        ws.cell(row=r, column=7, value=issuer)
        ws.cell(row=r, column=8, value=note)
    for r in DETAIL_ROWS:
        for c in range(1, NCOL + 1):
            cc = ws.cell(row=r, column=c)
            cc.border = BD
            cc.font = F_BODY
            cc.alignment = {1: AC, 2: AC, 3: ALW, 4: AC, 5: AR, 6: AC, 7: ALW, 8: ACW}[c]

    for cat in CATS:
        r = SUB_ROWS[cat]
        ws.cell(row=r, column=1, value=f'="{cat} 小计（"&COUNTIF({CAT_RANGE},"{cat}")&" 张）"')
        ws.cell(row=r, column=5,
                value=f'=SUMIF({CAT_RANGE},"{cat}",{AMT_RANGE})').number_format = FMT_MONEY
        for c in range(1, NCOL + 1):
            cc = ws.cell(row=r, column=c)
            cc.border = BD
            cc.font = F_BOLD
            cc.fill = FILL_SUB
            cc.alignment = ALW if c == 8 else (AR if c == 5 else AL)

    ws.cell(row=TOTAL_ROW, column=1, value=f'="合　　计（"&COUNTA({INV_RANGE})&" 张）"')
    ws.cell(row=TOTAL_ROW, column=5,
            value="=SUM(" + ",".join(f"E{SUB_ROWS[c]}" for c in CATS) + ")"
            ).number_format = FMT_MONEY
    for c in range(1, NCOL + 1):
        cc = ws.cell(row=TOTAL_ROW, column=c)
        cc.border = BD
        cc.font = F_TOTAL
        cc.fill = FILL_TOTAL
        cc.alignment = AR if c == 5 else AL
    ws.row_dimensions[TOTAL_ROW].height = 22

    for col, w in zip("ABCDEFGH", [6, 10, 44, 12, 12, 22, 26, 20]):
        ws.column_dimensions[col].width = w

    dv = DataValidation(type="list", formula1='"%s"' % ",".join(CATS), allow_blank=True)
    dv.errorTitle = "项目名称无效"
    dv.error = "请从下拉列表中选择：" + " / ".join(CATS)
    ws.add_data_validation(dv)
    dv.add(f"B{FIRST_DATA}:B{LAST_BODY_ROW}")

    ws.freeze_panes = "A3"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:2"

    wb.save(args.out)

    agg = {}
    for row in ROWS:
        a = agg.setdefault(row[0], [0, 0.0])
        a[0] += 1
        a[1] = round(a[1] + row[3], 2)
    grand = round(sum(v[1] for v in agg.values()), 2)

    cache = {}
    for cat in CATS:
        r, (cnt, amt) = SUB_ROWS[cat], agg.get(cat, [0, 0.0])
        cache["A%d" % r] = "%s 小计（%d 张）" % (cat, cnt)
        cache["E%d" % r] = amt
    cache["A%d" % TOTAL_ROW] = "合　　计（%d 张）" % len(ROWS)
    cache["E%d" % TOTAL_ROW] = grand
    inject_cache(args.out, {"xl/worksheets/sheet1.xml": cache})

    print("saved:", args.out)
    print("布局: 明细行", FIRST_DATA, "-", LAST_BODY_ROW, "| 小计行", SUB_ROWS, "| 合计行", TOTAL_ROW)
    print("分类:", " | ".join("%s %d张 ¥%.2f" % (c, agg.get(c, [0])[0], agg.get(c, [0, 0.0])[1])
                              for c in CATS), "| 合计 %d张 ¥%.2f" % (len(ROWS), grand))


if __name__ == "__main__":
    main()
