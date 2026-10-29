---
name: "兰台智报"
description: "智能数据报表与办公自动化平台（Windows）：通过 COM 自动化操控 WPS 表格，一条命令完成数据报表的汇总公式、区域格式、图表插图与 XLSX/PDF 双格式导出；并扩展 WPS 文字/演示、Zotero 学术文献管理、Photoshop/Illustrator 设计软件控制。使用前先按需读取对应子技能 SKILL.md。"
---

# 兰台智报

智能数据报表与办公自动化平台：让 AI Agent 通过 COM 自动化直接操控 Windows 上的办公与设计软件，
尤其擅长「数据报表与分析」——从 CSV 数据源一键生成带汇总公式、格式与图表的 WPS 表格报表，
并导出 XLSX/PDF。

## 技能组成

本平台聚合了 4 个子能力，按需读取对应子技能的 SKILL.md 后使用：

| 子技能 | SKILL.md 路径 | 能力 |
|--------|--------------|------|
| cli-anything-wps | `cli_anything/wps/skills/SKILL.md` | WPS 文字/表格/演示 CLI 操控，数据报表一键生成、JSON 数据驱动 PPT 自动生成 |
| cli-anything-zotero | `cli_anything/zotero/skills/SKILL.md` | Zotero 文献管理：检索、系统综述、论文起草、评审、期刊规范绘图 |
| cli-anything-photoshop | `photoshop-harness/skills/cli-anything-photoshop/SKILL.md` | Photoshop 位图处理：图层管理、选区、文本、导出 |
| illustrator-harness | `illustrator-harness/agent-harness/` | Illustrator 矢量图编程式创建与编辑 |

## 前置条件

- Windows 操作系统
- 目标软件已安装：WPS Office 12.0+（wps）/ Zotero（zotero）/ Adobe Photoshop / Adobe Illustrator
- Python 3.10+
- `pip install pywin32 matplotlib openpyxl`（数据报表）、`comtypes`（Photoshop）等对应依赖

## 顶层入口

```bash
# 安装（在仓库根目录）
pip install -e .
cli-anything-wps --help   # WPS（含数据报表命令）
cli-anything-zotero --help  # Zotero
```

## 核心场景：数据报表与分析（一键工作流）

```bash
cd report_demo
python build_report.py   # CSV 数据源 → 汇总公式/格式/图表 → XLSX + PDF
```

详细用法见 `report_demo/README.md` 与 `cli_anything/wps/skills/SKILL.md` 的「数据报表与分析」章节。

## 参考案例

- `report_demo/`：生产运行数据周报一键生成工作流（本平台核心案例，含 build 脚本、图表生成与双格式导出）
- `WPS/` 目录下含 16 个高校招生 PPT 制作案例（清华/北大/复旦/中科大等），含完整 build 脚本与 data.json，可直接复用为 JSON 驱动 PPT 工作流模板。
