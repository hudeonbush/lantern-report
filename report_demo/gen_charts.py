#!/usr/bin/env python3
"""生成数据周报图表（matplotlib）→ PNG 文件。

用法:
    python gen_charts.py

输出:
    charts/各装置产量与能耗对比.png   （柱状图）
    charts/一周产量与开工率趋势.png   （折线图）

说明:
    使用 matplotlib 生成图表（与 WPS 演示工作流的图表策略一致），
    再通过 calc image add / build_report.py 将 PNG 插入 WPS 表格。
"""

import csv
import os
import warnings

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# 中文字体（Windows SimHei；缺失时回退并抑制警告）
try:
    plt.rcParams["font.sans-serif"] = ["SimHei", "Microsoft YaHei", "Arial Unicode MS", "sans-serif"]
    plt.rcParams["axes.unicode_minus"] = False
    # 触发字体解析，若 SimHei 缺失则回退
    fig, ax = plt.subplots()
    ax.set_title("字体探测")
    plt.close(fig)
except Exception:
    pass
warnings.filterwarnings("ignore", category=UserWarning, module="matplotlib")

DEMO_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(DEMO_DIR, "data")
CHART_DIR = os.path.join(DEMO_DIR, "charts")

PRIMARY = "#1A3C8B"
SECONDARY = "#5B9BD5"
TOTAL = "#DCE6F1"


def load_csv(path: str) -> list:
    with open(path, "r", encoding="utf-8-sig") as f:
        return list(csv.reader(f))


def main() -> None:
    os.makedirs(CHART_DIR, exist_ok=True)

    # ── 图1：各装置产量与能耗对比（柱状图） ─────────────────
    weekly = load_csv(os.path.join(DATA_DIR, "周数据.csv"))
    header, rows = weekly[0], weekly[1:]
    names = [r[0] for r in rows]
    output_v = [float(r[1]) for r in rows]
    energy_v = [float(r[2]) for r in rows]

    fig, ax1 = plt.subplots(figsize=(9, 4.8), dpi=150)
    bars = ax1.bar(names, output_v, color=PRIMARY, width=0.45, label="产量(吨)")
    ax1.set_ylabel("产量(吨)", color=PRIMARY)
    ax1.tick_params(axis="y", labelcolor=PRIMARY)
    ax1.set_ylim(0, max(output_v) * 1.25)
    for b, v in zip(bars, output_v):
        ax1.text(b.get_x() + b.get_width() / 2, v + max(output_v) * 0.03,
                 f"{v:,.0f}", ha="center", fontsize=9)

    ax2 = ax1.twinx()
    ax2.plot(names, energy_v, color=SECONDARY, marker="o",
             linewidth=2, label="能耗(kgce/t)")
    ax2.set_ylabel("能耗(kgce/t)", color=SECONDARY)
    ax2.tick_params(axis="y", labelcolor=SECONDARY)
    ax2.set_ylim(0, max(energy_v) * 1.3)

    ax1.set_title("各装置产量与能耗对比", fontsize=13, pad=12)
    fig.tight_layout()
    fig.savefig(os.path.join(CHART_DIR, "各装置产量与能耗对比.png"),
                bbox_inches="tight")
    plt.close(fig)

    # ── 图2：一周产量与开工率趋势（折线图） ─────────────────
    trend = load_csv(os.path.join(DATA_DIR, "日趋势.csv"))
    _, trows = trend[0], trend[1:]
    dates = [r[0] for r in trows]
    daily = [float(r[1]) for r in trows]
    rate = [float(r[2]) for r in trows]

    fig, ax1 = plt.subplots(figsize=(9, 4.8), dpi=150)
    ax1.plot(dates, daily, color=PRIMARY, marker="o", linewidth=2,
             label="日产量(吨)")
    ax1.set_ylabel("日产量(吨)", color=PRIMARY)
    ax1.tick_params(axis="y", labelcolor=PRIMARY)
    ax1.set_ylim(0, max(daily) * 1.25)
    for i, v in enumerate(daily):
        ax1.text(i, v + max(daily) * 0.03, f"{v:,.0f}",
                 ha="center", fontsize=9)

    ax2 = ax1.twinx()
    ax2.plot(dates, rate, color=SECONDARY, marker="s", linewidth=2,
             linestyle="--", label="开工率(%)")
    ax2.set_ylabel("开工率(%)", color=SECONDARY)
    ax2.tick_params(axis="y", labelcolor=SECONDARY)
    ax2.set_ylim(0, 100)

    ax1.set_title("一周日产量与开工率趋势", fontsize=13, pad=12)
    fig.tight_layout()
    fig.savefig(os.path.join(CHART_DIR, "一周产量与开工率趋势.png"),
                bbox_inches="tight")
    plt.close(fig)

    print(f"[OK] 图表已生成: {CHART_DIR}")
    print("  - 各装置产量与能耗对比.png")
    print("  - 一周产量与开工率趋势.png")


if __name__ == "__main__":
    main()
