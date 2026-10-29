#!/usr/bin/env python3
"""生产运行数据周报自动生成 —— 兰台智报 WPS 表格自动化（数据报表与分析）

一键脚本：读取 CSV 数据源 → matplotlib 生成图表 PNG → 构建格式化报表
（数据 + 汇总公式 + 表头样式 + 图表图片）→ 导出 XLSX + PDF。

用法:
    python build_report.py

输入:
    data/周数据.csv   各装置产量/能耗/优级品率（表头行 + 数据行）
    data/日趋势.csv   一周每日产量/开工率

输出:
    charts/各装置产量与能耗对比.png      （柱状图）
    charts/一周产量与开工率趋势.png      （折线图）
    output/生产运行数据周报.xlsx
    output/生产运行数据周报.pdf

正式使用时：将 data/ 下的 CSV 替换为真实数据（表头保持不变），重新运行本脚本即可。
"""

import csv
import logging
import os
import subprocess
import sys

# 将仓库根目录加入模块搜索路径
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO_ROOT)

LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "run.log")
logging.basicConfig(filename=LOG_FILE, level=logging.INFO,
                    format="%(asctime)s %(message)s", force=True)

from cli_anything.wps.core import document as doc_mod
from cli_anything.wps.core import calc as calc_mod
from cli_anything.wps.core import export as export_mod

DEMO_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(DEMO_DIR, "data")
CHART_DIR = os.path.join(DEMO_DIR, "charts")
OUT_DIR = os.path.join(DEMO_DIR, "output")

# 报表配色（品牌蓝）
HEADER_FILL = "1A3C8B"   # 表头深蓝
HEADER_FONT = "FFFFFF"   # 表头白字
TOTAL_FILL = "DCE6F1"    # 汇总行浅蓝


def load_csv(path: str) -> list:
    """读取 CSV（支持 UTF-8 BOM）。"""
    with open(path, "r", encoding="utf-8-sig") as f:
        return list(csv.reader(f))


def build_project(weekly_rows: list, trend_rows: list) -> dict:
    """构建 Calc 项目：周数据页（表+柱状图）+ 趋势分析页（表+折线图）。"""
    proj = doc_mod.create_document(
        doc_type="calc", name="生产运行数据周报",
    )

    # ── Sheet 1: 周数据 ──────────────────────────────────────
    calc_mod.rename_sheet(proj, 0, "周数据")

    calc_mod.set_range(proj, "A1", weekly_rows)

    n_data = len(weekly_rows) - 1            # 数据行数
    last_row = n_data + 1                    # 数据最后一行（含表头）
    sum_row = last_row + 1                   # 汇总行

    # 汇总行：产量、能耗求和；优级品率取平均
    calc_mod.add_summary(proj, f"B2:C{last_row}", func="sum",
                         label="合计", label_ref=f"A{sum_row}")
    calc_mod.add_summary(proj, f"D2:D{last_row}", func="average", label=None)

    # 格式：表头加粗白字深蓝底、居中对齐、加边框
    calc_mod.format_range(proj, "A1", f"D1",
                          bold=True, fill_color=HEADER_FILL,
                          font_color=HEADER_FONT, align="center", border=True)
    # 数据区：居中、加边框
    calc_mod.format_range(proj, "A2", f"D{last_row}", align="center", border=True)
    # 汇总行：浅蓝底加粗
    calc_mod.format_range(proj, f"A{sum_row}", f"D{sum_row}",
                          bold=True, fill_color=TOTAL_FILL, align="center", border=True)
    # 数值格式
    calc_mod.format_range(proj, "B2", f"B{last_row}", number_format="#,##0")
    calc_mod.format_range(proj, "C2", f"C{last_row}", number_format="#,##0.00")
    calc_mod.format_range(proj, "D2", f"D{last_row}", number_format="#,##0.0")

    # 列宽
    calc_mod.set_col_width(proj, "A", 16)
    calc_mod.set_col_width(proj, "B", 14)
    calc_mod.set_col_width(proj, "C", 16)
    calc_mod.set_col_width(proj, "D", 16)

    # 图表（matplotlib PNG 插入，绕开 WPS Chart 对象后台死锁）
    calc_mod.add_image(
        proj, os.path.join(CHART_DIR, "各装置产量与能耗对比.png"),
        sheet=0, left=500, top=10, width=480, height=260,
    )

    # ── Sheet 2: 趋势分析 ────────────────────────────────────
    calc_mod.add_sheet(proj, "趋势分析")
    calc_mod.set_range(proj, "A1", trend_rows, sheet=1)

    t_last = len(trend_rows)
    calc_mod.format_range(proj, "A1", "C1", sheet=1,
                          bold=True, fill_color=HEADER_FILL,
                          font_color=HEADER_FONT, align="center", border=True)
    calc_mod.format_range(proj, "A2", f"C{t_last}", sheet=1,
                          align="center", border=True)
    calc_mod.format_range(proj, "B2", f"B{t_last}", sheet=1, number_format="#,##0")
    calc_mod.format_range(proj, "C2", f"C{t_last}", sheet=1, number_format="#,##0.0")
    calc_mod.set_col_width(proj, "A", 12, sheet=1)
    calc_mod.set_col_width(proj, "B", 16, sheet=1)
    calc_mod.set_col_width(proj, "C", 14, sheet=1)

    calc_mod.add_image(
        proj, os.path.join(CHART_DIR, "一周产量与开工率趋势.png"),
        sheet=1, left=500, top=10, width=480, height=260,
    )

    return proj


def main() -> None:
    # WPS COM 前置清理：正在运行的 WPS 实例的模态弹窗会阻塞 COM 调用
    # （详见 SKILL.md「执行前清理」经验）。请先保存并关闭你正在编辑的 WPS 文档。
    from cli_anything.wps.utils.wps_backend import kill_all_wps_processes
    kill_all_wps_processes()
    logging.info("WPS 进程已清理")

    # 1. 生成图表 PNG
    subprocess.run([sys.executable, os.path.join(DEMO_DIR, "gen_charts.py")],
                   check=True)
    logging.info("图表 PNG 已生成")

    # 2. 构建并导出
    weekly = load_csv(os.path.join(DATA_DIR, "周数据.csv"))
    trend = load_csv(os.path.join(DATA_DIR, "日趋势.csv"))
    proj = build_project(weekly, trend)
    logging.info("项目构建完成，开始导出")

    os.makedirs(OUT_DIR, exist_ok=True)

    # 导出 XLSX
    xlsx = os.path.join(OUT_DIR, "生产运行数据周报.xlsx")
    result = export_mod.export(proj, xlsx, preset="xlsx", overwrite=True)
    print(f"[OK] XLSX → {result['output']} ({result['file_size']:,} 字节, "
          f"WPS {result.get('wps_version')})")
    logging.info("XLSX 导出成功: %s", result['file_size'])

    # 导出 PDF
    pdf = os.path.join(OUT_DIR, "生产运行数据周报.pdf")
    result = export_mod.export(proj, pdf, preset="pdf-calc", overwrite=True)
    print(f"[OK] PDF  → {result['output']} ({result['file_size']:,} 字节)")
    logging.info("PDF 导出成功: %s", result['file_size'])


if __name__ == "__main__":
    main()
