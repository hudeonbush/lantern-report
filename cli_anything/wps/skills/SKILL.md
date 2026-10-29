---
name: "cli-anything-wps"
description: "WPS Office CLI — JSON数据驱动PPT自动生成 + 命令行操控WPS文字/表格/演示文稿"
---

# cli-anything-wps

通过 COM 自动化接口操控 WPS Office 的命令行工具。支持 WPS 文字（Writer）、WPS 表格（Calc）和 WPS 演示（Impress）。

## 前置条件

- **Windows 操作系统**
- **WPS Office**（已安装，版本 12.0+）
- **Python 3.10+**
- **pywin32**：`pip install pywin32`

## 安装

```bash
pip install -e .
cli-anything-wps --help
```

---

## JSON 数据驱动 PPT 自动生成 ⭐ 生产环境核心工作流

**这是最常用的模式**：搜数据 → 生成 matplotlib 图表 → 写 JSON → WPS COM 一键生成。

### 完整流程（6步）

```
1. 提取模板背景: 模版.pptx → template_bg.png (960x540)
2. WebSearch 搜索数据: 招生分数线/学科排名/招生计划/科研数据
3. matplotlib 生成图表: 柱状图/饼图/折线图/横向柱/气泡图
4. 编写 data.json: 12-15页 elements[] 编排（标题≤4字 + 间距≥24pt）
5. python build_xxx.py: WPS COM 逐页构建
6. 输出: PPTX + PDF 双格式
```

### 标准页序

```
S1 封面 → S2 目录 → S3-S10 内容页(图+表+卡片) → S11 总结卡片 → S12 致谢
```

### 元素类型路由

| type | 用途 |
|------|------|
| `text` | 文本框 |
| `image` | 图片(含matplotlib图表) |
| `table` | 表格 |
| `cards_2x3` | 2行x3列彩色卡片 |
| `cards_1x4_info` | 4列数字统计卡 |
| `card_list_wide` | 目录编号列表 |
| `tagline_bar` | 页面底部总结条 |

### 关键约束 ⚠️

- **标题与内容间距**：标题 y=14 h=40(结束于y≈54)，第一个内容元素**必须**起始于 y≥76-78(≥24pt gap)
- **标题最多4字**，居中，SimHei 40-44pt，品牌色，**无装饰线**
- **JSON中文引号**：文本内引用用「」代替 `""`
- **所有元素不出画布**：960×540，y+h≤518
- **WPS COM**：`Fill.ForeColor.RGB`，`SaveAs(path, 32)`导出PDF
- **执行前清理**：`taskkill //F //IM wps.exe //T`

### 参考案例

[WPS/](../WPS/) 目录下16个项目：

| 项目 | 页数 | 主题 |
|------|------|------|
| 清华协和/兰州大学/同济医学院/哈工大/重庆大学/南华大学 | 12页 | 各校招生 |
| 中山大学/中科大/国科大 | 15页 | 名校招生 |
| 复旦大学 | 12页 | 新工科+医学院 |
| 南科大肿瘤医院 | 12页 | 联合培养硕博 |
| 北大/清华/南科大/华中科大/浙大城市学院 | 9-14页 | 各校介绍 |

```bash
# 完整执行示例
python -c "import zipfile; ..."  # 提取模板
python gen_charts.py              # 生成图表
python -c "import json; json.load(open('data.json','r',encoding='utf-8'))"  # 验证JSON
taskkill //F //IM wps.exe //T
python build_xxx.py               # 构建PPTX+PDF
```

---

## 数据报表与分析（Calc 自动化）⭐ 参赛主线工作流

**核心模式**：CSV/数据源 → 批量写入 → 汇总公式 → 格式化 → 图表 → 一键导出 XLSX/PDF。

### 一键工作流（推荐）

```bash
python report_demo/build_report.py   # 读取 data/*.csv → 生成周报 xlsx+pdf
```

完整演示见 `report_demo/README.md`：数据源（`data/周数据.csv`、`data/日趋势.csv`）
替换为真实数据后重新运行即可，表头保持 `装置,产量(吨),能耗(kgce/t),优级品率(%)` 等结构。

### Calc 命令速查（v1.1 新增）

| 命令 | 用途 | 示例 |
|------|------|------|
| `calc set-range` | 批量写入二维数据 | `set-range A1 --data '[[...]]'` |
| `calc summary` | 汇总行公式（sum/avg/count/max/min） | `summary --data-range B2:C7 --func sum --label 合计 --label-ref A8` |
| `calc format-range` | 表头加粗/填充色/边框/数字格式/对齐 | `format-range --start A1 --end D1 --bold --fill 1A3C8B --font-color FFFFFF --align center --border` |
| `calc col-width` | 整列列宽 | `col-width A 16` |
| `calc chart add` | 图表（column/bar/line/pie/area/scatter/doughnut） | `chart add --type column --data-range A1:D7 --title "产量对比"` |
| `calc image add` | 插入图片（matplotlib 图表 PNG 等，**推荐用于报表图表**） | `image add charts/图.png --sheet 0 --left 500 --top 10 --width 480 --height 260` |
| `calc merge-cells` | 合并单元格 | `merge-cells A1 D1` |
| `export render` | 导出 xlsx/csv/pdf 等 | `export render 周报.xlsx --preset xlsx --overwrite` |

### 报表生成标准步骤（Agent 使用）

```
1. 读取数据源（CSV/表格/接口）→ 二维数组
2. document new --type calc → rename-sheet 数据页
3. set-range 写入表头+数据 → summary 生成汇总行（数值列用 SUM，比率列用 AVERAGE）
4. format-range：表头深蓝白字加粗居中、数据区边框居中、汇总行浅蓝、数字格式（#,##0.00）
5. set-col-width 调列宽
6. 图表：优先 matplotlib 生成 PNG（report_demo/gen_charts.py 模板）→ calc image add 插入表格右侧
7. export render 导出 XLSX + PDF（--overwrite；PDF 每工作表自动横向单页）
```

### 图表类型路由

| 场景 | 图表类型 |
|------|---------|
| 多装置/多项目对比 | `column`（柱状图） |
| 时间序列趋势 | `line`（折线图） |
| 结构占比 | `pie`（饼图） |
| 排名分布 | `bar`（横向条形图） |

### 质量要点 ⚠️

- 表头：加粗 + 深蓝底（`1A3C8B`）+ 白字（`FFFFFF`）+ 居中 + 边框
- 汇总行：浅蓝底（`DCE6F1`）+ 加粗；数值列 `=SUM()`，比率列用 `=AVERAGE()`
- 数字格式：产量 `#,##0`，能耗 `#,##0.00`，比率 `#,##0.0`
- 颜色格式为 `RRGGBB` 十六进制（内部自动转 COM BGR）
- **图表策略**：报表图表推荐「matplotlib 生成 PNG → `calc image add` 插入」
  （与 PPT 工作流一致）。WPS 表格的 Chart 对象在后台模式（Visible=False）下
  与单元格格式（加粗/填充/边框）共存时保存不稳定（实测 E_FAIL/挂起），
  PNG 插图方案全链路稳定可复现。`calc chart add` 仍保留用于简单场景
  （数据源仅数字格式或无格式时可用）。
- **导出前清理**：`kill_all_wps_processes()`（wps_backend 提供）——正在运行的
  WPS 实例的模态弹窗会阻塞 COM 调用（RPC_E_CALL_REJECTED），必须先结束。
- **初始化等待**：find_wps 已内置 COM 初始化探测（Version 轮询 + 1s 等待），
  并设置 `DisplayAlerts=False` 抑制文档恢复对话框等模态弹窗。
- **多工作表顺序**：WPS 的 `Worksheets.Add()` 默认插到最前且 `Add(After=)`
  不可靠，export 渲染采用**逆序填充**保证最终顺序与项目定义一致。
- **PDF 排版**：export 对每个工作表自动设置横向 + `FitToPagesWide=1/Tall=1`，
  保证 PDF 每表一页、无空白。

---

## CLI 命令结构

```
wps
├── document          # 文档管理: new/open/save/info
├── writer            # 文字: add-paragraph/heading/list/table/image
├── calc              # 表格: set-cell/get-cell/set-range/merge-cells
├── impress           # 演示: add-slide/remove-slide/set-content/add-element
├── style             # 样式: create/modify/list/apply/remove
├── preset            # 设计预设: list/info/apply
├── export            # 导出: render output.pptx -p pptx
└── session           # 会话: status/undo/redo/history
```

### 使用示例

```bash
# 创建文档
cli-anything-wps document new --type writer --name "报告" -o report.json
cli-anything-wps --project report.json writer add-heading -t "年度报告" -l 1
cli-anything-wps --project report.json export render report.docx -p docx

# JSON模式（Agent使用）
cli-anything-wps --json document new --type writer --name "test"
```

### 设计预设

| 预设 | 配色 | 适用场景 |
|------|------|---------|
| academic | 深蓝#1A3C8B | 学术答辩/基金申请 |
| consultant | 深蓝#003366+亮青 | 咨询报告 |
| business | 商务蓝#005294 | 会议汇报/课件 |
| tech | 近黑#0F1423 | 科技/AI/数据报告 |

### 质量审查（5维度）

| 维度 | 检查项 | 阈值 |
|------|--------|------|
| visual | 字体层级/对比度/留白 | 70分 |
| pedagogy | 叙事弧/预备知识 | 75分 |
| proofreading | 拼写/术语/溢出 | 80分 |
| parity | PPTX/PDF一致性 | 85分 |
| substance | 数据准确性/引用 | 90分 |
