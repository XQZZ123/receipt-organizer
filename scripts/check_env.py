# -*- coding: utf-8 -*-
"""环境自检：依赖、项目结构、配置。

换到新机器/新项目后先跑这个，缺什么它会说清楚（依赖可直接自动装）。

用法：
    python check_env.py --root <项目根目录> [--install]
"""
import argparse
import importlib.util
import os
import subprocess
import sys

REQUIRED = [("pymupdf", "pymupdf>=1.24"), ("openpyxl", "openpyxl>=3.1")]
LEDGER_DIR = ".invoice-ledger"


def have(mod):
    return importlib.util.find_spec(mod) is not None


def ensure(mod, spec, install):
    if have(mod):
        return True, "已安装"
    if not install:
        return False, f"缺失，运行 `pip install {spec}` 或加 --install"
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", spec])
        return have(mod), "已自动安装" if have(mod) else "安装失败"
    except subprocess.CalledProcessError as exc:
        return False, f"安装失败：{exc}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True, help="项目根目录（发票月份目录的上级）")
    ap.add_argument("--install", action="store_true", help="缺失依赖时自动 pip install")
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    print(f"Python: {sys.version.split()[0]}  ({sys.executable})")
    print(f"项目根: {root}\n")

    ok = True
    print("依赖：")
    for mod, spec in REQUIRED:
        good, msg = ensure(mod, spec, args.install)
        ok &= good
        print(f"  [{'OK' if good else 'MISSING'}] {mod:<10} {msg}")

    print("\n项目结构：")
    if not os.path.isdir(root):
        print(f"  [FAIL] 根目录不存在：{root}")
        sys.exit(1)
    months = sorted(d for d in os.listdir(root)
                    if os.path.isdir(os.path.join(root, d)) and d.isdigit() and len(d) == 6)
    print(f"  月份目录：{', '.join(months) if months else '无（首次运行会自动建）'}")

    ledger = os.path.join(root, LEDGER_DIR)
    cfg_path = os.path.join(ledger, "config.json")
    if os.path.exists(cfg_path):
        import json
        with open(cfg_path, encoding="utf-8") as fh:
            cfg = json.load(fh)
        buyer = (cfg.get("buyer") or "").strip()
        if buyer:
            print(f"  [OK] 公司抬头：{buyer}")
        else:
            ok = False
            print(f"  [WARN] 未填公司抬头 → 请编辑 {cfg_path} 的 \"buyer\"")
    else:
        os.makedirs(ledger, exist_ok=True)
        import json
        with open(cfg_path, "w", encoding="utf-8") as fh:
            json.dump({"_说明": "buyer 填本公司发票抬头全称（与票面「购买方名称」完全一致）；cats 是费用科目分区顺序。",
                       "buyer": "",
                       "cats": ["交通费", "住宿费", "餐饮费", "其他费用"]},
                      fh, ensure_ascii=False, indent=2)
        ok = False
        print(f"  [新建] 已生成配置模板：{cfg_path}")
        print(f"         → 填上 \"buyer\" 后重跑本脚本")

    excl = os.path.join(ledger, "excluded.json")
    print(f"  排除名单：{'已有' if os.path.exists(excl) else '无（用到时会建）'}")

    print("\n结论：" + ("环境就绪" if ok else "按上面标出的项目补齐后重跑"))
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
