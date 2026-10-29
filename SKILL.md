---
name: "Lantern"
description: "Lantern（灯台）——以自动设计 PPT 为核心的办公自动化平台（Windows）：通过 COM 自动化操控 WPS 演示，一份 JSON 数据即定义整套幻灯片，一条命令自动生成成品 PPTX 与 PDF，并驱动 WPS 演示播放与导出，全程无需手工排版。附带 WPS 表格数据报表能力（CSV 一键生成带汇总公式与图表插图的 XLSX/PDF），并扩展 Zotero 学术文献管理、Photoshop/Illustrator 设计软件控制。使用前先按需读取对应子技能 SKILL.md。"
---

# Lantern（灯台）

以**自动设计 PPT**为核心的办公自动化平台：让 AI Agent 通过 COM 自动化直接操控 Windows 上的办公与设计软件。
它的核心主张是**全自动化**——把「设计一套 PPT」从手工排版变为一条命令：

```text
一份 JSON 数据（封面 / 图文 / 表格 / 红条） → 一条命令 → 成品 PPTX + PDF（可播放、可导出）
```

改文案、改预算、改日程，只需改 JSON 数字重新运行，即可一键得到新版，**全程零手工排版、可复现**。

## 技能组成

本平台聚合了 4 个子能力，按需读取对应子技能的 SKILL.md 后使用：

| 子技能 | SKILL.md 路径 | 能力 |
|--------|--------------|------|
| cli-anything-wps | `cli_anything/wps/skills/SKILL.md` | WPS 文字/表格/演示 CLI 操控，JSON 数据驱动 PPT 自动设计、数据报表一键生成 |
| cli-anything-zotero | `cli_anything/zotero/skills/SKILL.md` | Zotero 文献管理：检索、系统综述、论文起草、评审、期刊规范绘图 |
| cli-anything-photoshop | `photoshop-harness/skills/cli-anything-photoshop/SKILL.md` | Photoshop 位图处理：图层管理、选区、文本、导出 |
| illustrator-harness | `illustrator-harness/agent-harness/` | Illustrator 矢量图编程式创建与编辑 |

## 前置条件

- Windows 操作系统
- 目标软件已安装：WPS Office 12.0+（wps）/ Zotero（zotero）/ Adobe Photoshop / Adobe Illustrator
- Python 3.10+
- `pip install pywin32>=306 Pillow>=10.0 pymupdf>=1.23 matplotlib openpyxl`（PPT 自动设计 + 数据报表）、`comtypes`（Photoshop）等对应依赖

## 顶层入口

```bash
# 安装（在仓库根目录）
pip install -e .
cli-anything-wps --help   # WPS（含 PPT 自动设计命令）
cli-anything-zotero --help  # Zotero
```

## 核心场景：自动设计 PPT（一键工作流）

```bash
cd WPS/ppt-engine/examples
python build_from_json.py --json lantern/示例1_徒步策划案_数据源.json --out lantern/out --name 示例1_徒步策划案_精简
# JSON 数据 → 成品 PPTX + PDF（自动校验页数与溢出）
```

- 数据源：`lantern/示例1_徒步策划案_数据源.json`（3 页徒步团建活动策划案，演示案例，参数可调整）
- 生成器：`WPS/ppt-engine/examples/build_from_json.py`（支持封面/图文/表格/红条等元素，多列表格自动渲染）
- 详细用法见 `cli_anything/wps/skills/SKILL.md` 的「自动设计 PPT」章节。

## 参考案例

- `WPS/ppt-engine/examples/lantern/`：**宁波金发新材料有限公司徒步团建活动策划案**（3 页，含封面/行程表格/预算表格，Lantern 全自动生成，改 JSON 即得新版）
- `WPS/ppt-engine/examples/`：16 个高校招生 PPT 案例（清华/北大/复旦/中科大等），含完整 build 脚本与 data.json，可直接复用为 JSON 驱动 PPT 模板
- `report_demo/`：生产运行数据周报一键生成工作流（数据报表附带能力，CSV → XLSX + PDF）
