# 报销整理 skill — 使用说明

把每月从邮箱收到的发票，整理成一份能直接交财务的 Excel。
一个月一条流水线：**邮箱取票 → 改名归位 → 查重 → 汇总表 → 校验**。

做好了的东西：

- 发票按 `YYYYMM` 分月归档，文件名统一成 `日期-做什么-金额元.pdf`
- 自动跨月查重，防止把上个月报过的票再报一遍
- 出一张按费用科目分块、带小计的 Excel
- 交付前跑 7 项校验，任何一项不过就不让你交

---

## 一、装（三步）

### 1. 放对位置

把整个 `receipt-organizer` 文件夹放到**你项目的** `.workbuddy/skills/` 下。
用 git 装就是一行，目录名别改 —— skill 靠这个名字加载：

```bash
cd 你的项目/.workbuddy/skills
git clone <仓库地址> receipt-organizer
```

没装 git 就下载 zip 解压，把里面的 `receipt-organizer` 整个挪过去。装完长这样：

```
你的项目/
├─ 202609/            ← 月份目录，脚本自己建
├─ 202610/
└─ .workbuddy/
   └─ skills/
      └─ receipt-organizer/     ← 就放这里
         ├─ SKILL.md
         ├─ README.md
         └─ scripts/
```

项目级 skill 只在这个项目里生效 —— 换个项目要再放一份。

### 2. 查环境

```bash
python .workbuddy/skills/receipt-organizer/scripts/check_env.py --root . --install
```

`--install` 会把缺的依赖自动装上（就两个：`pymupdf` 读 PDF、`openpyxl` 写 Excel）。

如果 `python` 命令找不到，用你装了 Python 的那个解释器路径。WorkBuddy 自带的通常在：

- Windows：`~/.workbuddy/binaries/python/envs/default/Scripts/python.exe`
- macOS / Linux：`~/.workbuddy/binaries/python/envs/default/bin/python`

### 3. 填公司抬头（必须）

上一步会生成 `.invoice-ledger/config.json`，把 `buyer` 填成**你们公司发票上的购买方全称**，一字不差：

```json
{
  "buyer": "某某科技有限公司",
  "cats": ["交通费", "住宿费", "餐饮费", "其他费用"]
}
```

**不填会怎样**：发票上只有「购买方」和「销售方」两个公司名，脚本没法判断哪个是卖方，开票单位就只能留空。填了才能自动抽对。
填错会怎样：每张票都会提示「票面未出现抬头」，你立刻就知道配错了，不会静默出错。

填完再跑一次 `check_env.py`，看到「环境就绪」就行。

> 这一步不顺的话，先确认抬头是**发票上印的那个全称**（含「有限公司」「股份有限公司」这些字），别填简称。

### 4. 邮箱连接器

skill 本身不连邮箱 —— 它借用 WorkBuddy 的邮箱连接器（QQ 邮箱、Gmail、Outlook 等），
**你得先在连接器里把自己的邮箱授权好**。没有连接器，第一步「取票」就没法做。

没有连接器也能用：手动把发票 PDF 丢进月份目录，从「改名归位」那步开始跑。

---

## 二、用

### 在 WorkBuddy 里直接说

装好之后，直接跟它说：

- 「整理这个月的发票」
- 「把邮箱里的发票下载下来，做报销表」
- 「查一下有没有重复报销的」

skill 会自动加载，按流程跑。

### 或手动一条条跑

```bash
export PYTHONIOENCODING=utf-8
ROOT="."                                    # 项目根目录
SK="$ROOT/.workbuddy/skills/receipt-organizer/scripts"
PY="python"
MONTH="$(date +%Y%m)"                       # 当月，如 202610

# 改名归位（先 dry-run 看计划，确认后再去掉 --dry-run 真跑）
"$PY" "$SK/organize.py" --root "$ROOT" --month "$MONTH" --dry-run

# 扫描 + 查重
"$PY" "$SK/scan_invoices.py" --root "$ROOT" --month "$MONTH"

# 建表（先把明细写进 rows.json，格式见 SKILL.md）
"$PY" "$SK/build_xlsx.py" --rows "$ROOT/$MONTH/_raw/rows.json" \
     --out "$ROOT/$MONTH/$(date +%Y)年$((10#$(date +%m)))月发票整理表.xlsx"

# 交付前校验
"$PY" "$SK/check_delivery.py" --xlsx "$ROOT/$MONTH/xxx.xlsx" --month-dir "$ROOT/$MONTH"
```

`rows.json` 要你自己看票面写（每张票的「项目内容」得写清楚是哪个行程/哪次加油），
脚本不猜这个 —— 猜出来的内容财务看不懂。

---

## 三、你会得到什么

```
你的项目/
├─ 202603/
│  ├─ 2026-3-6-高铁改签费-北京南-上海虹桥-20元.pdf
│  ├─ 2026-3-12-加油-128.4元.pdf
│  ├─ 2026-3-14-打车-36元.pdf
│  └─ 2026年3月发票整理表.xlsx
└─ .invoice-ledger/
   ├─ config.json      ← 抬头、科目顺序
   └─ excluded.json    ← 排除了哪些票（比如给同事报了）
```

Excel 长这样（一张表，不分两张 sheet）：

| 序号 | 项目名称 | 项目内容 | 日期 | 金额（元） | 发票号码 | 开票单位 | 备注 |
|---|---|---|---|---|---|---|---|
| 1 | 交通费 | G1234 北京南 → 上海虹桥，改签费 | 2026-3-6 | ¥20.00 | 12345678… | 中国铁路（北京南站） | 无 |
| 2 | 交通费 | 92号车用汽油 40 升 | 2026-3-12 | ¥128.40 | 09876543… | 某某石油销售有限公司 | 无 |
| | **交通费 小计（2 张）** | | | **¥148.40** | | | |
| | **合　　计（2 张）** | | | **¥148.40** | | | |

日期取**出行/消费日期**，不是开票日期。上个月遗留、这个月才开的票照收，日期如实填。

---

## 四、改成你们公司的口径

不用改代码，动这几处：

| 想改什么 | 改哪里 |
|---|---|
| 公司抬头 | `.invoice-ledger/config.json` → `buyer` |
| 费用科目 | `config.json` 的 `cats` + `scan_invoices.py` 的 `CATEGORY_RULES`（两处要一致） |
| 哪些票面词算哪类 | `scan_invoices.py` 的 `CATEGORY_RULES` |
| 文件名里的「做什么」 | `organize.py` 的 `WHAT_BY_KIND` |
| 表的列、配色、备注默认值 | `build_xlsx.py` 顶部常量 |
| 邮箱换成 Gmail/Outlook | Step 2 那三个调用换掉即可，文件落到 `_raw/` 就不用动脚本 |

---

## 五、几个坑（都实际踩过）

**发票 PDF 里藏着异体字。** 「长春」会写成 `⻓春`（U+2ED3，康熙部首），「金拱门」写成 `⾦拱⻔`。
脚本里已经用 NFKC 归一化处理了。你要是自己写正则匹配票面文字，记得先做这一步，否则永远匹配不上。

**不同开票平台的票面排版完全不一样。** 高德那几家（享道/如祺/曹操）连「发票号码」这几个字的位置都不同，
所以只能按特征匹配（20 位连续数字＝发票号码），不能按位置取。

**高德一封邮件能塞 16 个附件**：4 张票 × (pdf+ofd+xml) + 4 份行程单。
只取 pdf，且文件名含「行程单」的直接跳过，那不是发票。

**12306 发的是 zip**，里面才是 pdf/ofd/xml。脚本会自动解开只取 PDF。

**改签费和改签后的票价是两张票，都要收。** 前者票面写「改签费:」只有手续费，后者写「票价:」是全额。
别当成重复给删了 —— 文件名里也做了区分（`高铁改签费-…` vs `高铁票-…`）。

**某张票给同事报了怎么办？** 别只从表里删那一行。
PDF 还在月份目录里，下次整理会被当成「还没录的新票」又加回来，那就真重复报销了。
要登记到 `.invoice-ledger/excluded.json`，脚本会把票挪进 `_excluded/` 并永久跳过它。

**Excel 开着的时候别写表。** 写入可能"成功"，但你窗口里是旧数据，一保存就把改动覆盖回去了。
先关掉，或者写完关掉重开。

---

## 六、限制

- **只认 PDF**。图片型发票（拍照/扫描的）没有文本层，这些脚本读不出来，得走 OCR。
- **`rows.json` 的项目内容要人写**。脚本能抽出金额、日期、发票号码、开票单位，但「这趟是去哪儿」只能你看票面写。
- **依赖邮箱连接器**。没有就手动放 PDF，从改名那步开始。
- **一个项目一份**。项目级 skill 不跨项目共享，多个项目要各放一份。
