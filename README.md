# 兰台智报

<h3 align="center">智能数据报表与办公自动化平台 —— 一键生成数据报表，成册归档</h3>

<p align="center">
  <img src="https://img.shields.io/badge/平台-Windows-blue?logo=windows" alt="Platform">
  <img src="https://img.shields.io/badge/Python-3.10+-blue?logo=python" alt="Python">
  <img src="https://img.shields.io/badge/协议-MIT-green" alt="License">
  <img src="https://img.shields.io/badge/WPS-Office%2012.0+-red" alt="WPS">
</p>

---

**兰台智报**（Lantai Report）是面向「数据报表与分析」场景的智能报表平台：
通过 Windows COM 自动化操控 WPS 表格，一条命令完成 **CSV 数据源 → 汇总公式 → 区域格式 → 图表插图 → XLSX/PDF 双格式导出** 的完整报表生成链路，
并扩展 WPS 文字/演示、Zotero 学术文献管理、Photoshop/Illustrator 设计软件控制。

> 「兰台」典出古代官修档案之府，寓意数据报表经专业整理、成册归档；
> 「智报」即智能报表 —— 让 AI Agent 替你完成重复的报表制作。

## 核心场景：数据报表与分析（一键工作流）

```bash
cd report_demo
pip install pywin32 matplotlib openpyxl
python build_report.py
```

输出（约 6–10 秒）：

- `output/生产运行数据周报.xlsx` —— 双工作表（周数据 + 趋势分析），
  含汇总公式（`=SUM()`/`=AVERAGE()`）、统一表头样式、图表插图
- `output/生产运行数据周报.pdf` —— 每工作表一页，横向自适应，排版紧凑

数据源为 `data/*.csv`，替换真实数据（表头保持不变）重新运行即可复现。

## 平台能力

| 模块 | 能力 |
|------|------|
| `cli_anything/wps`（核心） | WPS 文字/表格/演示 CLI 操控：47 命令 + 数据报表能力（汇总公式、区域格式、列宽、图表插图、XLSX/PDF 导出） |
| `cli_anything/zotero` | Zotero 学术研究智能体：27 个 Skill（检索/综述/论文/审稿/期刊图表） |
| `illustrator-harness` | Adobe Illustrator 矢量图编程式创建与编辑 |
| `photoshop-harness` | Adobe Photoshop 位图处理：图层/选区/文本/滤镜/导出 |
| `WPS/` | 16 个高校招生 PPT 制作案例（JSON 数据驱动），含 ppt-engine 通用 PPT 引擎 |

## 数据报表能力详解（本次新增）

- **`calc.add_summary`**：自动生成 `SUM/AVERAGE/COUNT/MAX/MIN` 汇总行，公式范围由数据行数推导
- **`calc.format_range`**：加粗/填充色/字体色/边框/数字格式/对齐，RRGGBB 色值自动转换
- **`calc.set_col_width`**：列宽设置
- **`calc.add_chart` / `calc.add_image`**：7 种图表类型；推荐 matplotlib 生成 PNG 后插图（稳定方案）
- **`export`**：多工作表顺序保证、每表「横向 + 一页自适应」打印、XLSX/PDF 双导出

详细命令见 `cli_anything/wps/skills/SKILL.md`，一键工作流见 `report_demo/README.md`。

## 快速上手

```bash
# 数据报表（核心场景）
cd report_demo && python build_report.py

# WPS 办公
cli-anything-wps document new --type impress --name "演示"
cli-anything-wps export render output.pptx -p pptx

# Zotero 学术
cli-anything-zotero skills pipeline thesis

# Illustrator / Photoshop 设计
cli-anything-illustrator project new logo.ai -w 500 -h 500
cli-anything-photoshop project new banner.psd -w 1920 -h 1080
```

## 系统要求

- Windows 10/11
- WPS Office 12.0+（数据报表）/ Zotero 7+ / Adobe Illustrator 2023+ / Adobe Photoshop 2023+（可选）
- Python 3.10+ + pywin32 + matplotlib + openpyxl

## 许可证

MIT（见 [LICENSE](LICENSE)）
