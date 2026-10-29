# 生产运行数据周报自动生成（数据报表与分析演示工作流）

本目录是兰台智报在「数据报表与分析」场景的**可复现演示工作流**：
从 CSV 数据源一键生成格式化周报（数据页 + 趋势页 + 汇总公式 + 图表），并导出 XLSX 与 PDF。

## 痛点对应

日常工作中，生产运行数据周报/月报需要把各装置产量、能耗、优级品率等数据
手工复制进 Excel，再逐个加公式、插图表、调格式、导出，每次约 1–2 小时且易出错。
本工作流将整个过程固化为一条命令，从 CSV 到成品报表全自动。

## 运行前提

- Windows + WPS Office 12.0+（已安装，COM 自动化可用）
- Python 3.10+（实测 3.13）
- 依赖：`pip install pywin32 matplotlib openpyxl`
- 技能已在豆包工作台加载（或直接把 `cli_anything/` 加入 `PYTHONPATH`）

## 一键运行

```bash
cd report_demo
python build_report.py
```

输出：
- `charts/各装置产量与能耗对比.png`、`charts/一周产量与开工率趋势.png`
- `output/生产运行数据周报.xlsx`
- `output/生产运行数据周报.pdf`（每工作表一页，横向自适应）

## 输入数据源（正式使用时替换为真实数据）

| 文件 | 内容 |
|------|------|
| `data/周数据.csv` | 表头 `装置,产量(吨),能耗(kgce/t),优级品率(%)` + 各装置数据行 |
| `data/日趋势.csv` | 表头 `日期,日产量(吨),开工率(%)` + 一周每日数据行 |

> 当前 CSV 为**示例数据**，仅用于演示完整复现链路；正式使用时将两文件内容
> 替换为真实生产数据（表头保持不变）后重新运行即可。

## 脚本做了什么（提示词/操作步骤对应报告）

1. **前置清理**：`kill_all_wps_processes()` 结束正在运行的 WPS 实例——
   其模态弹窗会阻塞 COM 调用（重要经验，见 SKILL.md）。
2. `gen_charts.py`：用 matplotlib 生成两张图表 PNG
   （柱状图：各装置产量与能耗对比；折线图：一周日产量与开工率趋势）。
   图表先用 Python 生成图片、再插入表格——规避 WPS 表格 Chart 对象在
   后台模式下与单元格格式共存时的保存不稳定问题（已验证）。
3. `build_project()` 构建 Calc 项目（`document.create_document(type="calc")`）
4. **Sheet1 周数据**：
   - `calc.set_range` 批量写入表头与数据
   - `calc.add_summary` 生成汇总行：产量/能耗 `=SUM(...)`，优级品率 `=AVERAGE(...)`
   - `calc.format_range` 表头加粗白字深蓝底（1A3C8B/FFFFFF）、数据区居中加边框、
     汇总行浅蓝底（DCE6F1）、数值格式
   - `calc.set_col_width` 设置列宽
   - `calc.add_image` 插入柱状图 PNG（表格右侧）
5. **Sheet2 趋势分析**：同样格式 + 折线图 PNG
6. `export.export` 依次导出 XLSX 与 PDF（WPS COM 自动化，后台运行；
   每工作表自动横向 + 一页自适应，PDF 排版紧凑）

## 对应的 CLI 命令（等价于脚本内调用）

```bash
# 图表生成（独立步骤，可单独运行）
python gen_charts.py

# 手工等价命令（在 REPL / 项目模式下）
cli-anything-wps calc set-range A1 --data '[[...]]'
cli-anything-wps calc summary --data-range B2:C7 --func sum --label 合计 --label-ref A8
cli-anything-wps calc summary --data-range D2:D7 --func average --label none
cli-anything-wps calc format-range --start A1 --end D1 --bold --fill 1A3C8B --font-color FFFFFF --align center --border
cli-anything-wps calc col-width A --width 16
cli-anything-wps calc image add charts/各装置产量与能耗对比.png --sheet 0 --left 500 --top 10 --width 480 --height 260
cli-anything-wps export render output/生产运行数据周报.xlsx --preset xlsx --overwrite
cli-anything-wps export render output/生产运行数据周报.pdf --preset pdf-calc --overwrite
```

## 正式提交前核对

- 将 `data/*.csv` 替换为**真实工作数据**（表头保持不变）；
- 按部门实际口径调整汇总列（如能耗取均值而非求和）；
- 重新运行 `python build_report.py`，检查 XLSX 打开正常、PDF 每表一页。
