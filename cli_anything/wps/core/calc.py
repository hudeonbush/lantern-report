"""WPS CLI - Calc（电子表格）命令实现。"""

import os
from typing import Dict, Any, Optional, List
import re


def _ensure_calc(project: Dict[str, Any]) -> None:
    """确保项目是 Calc 类型。"""
    if project.get("type") != "calc":
        raise ValueError("当前文档不是 Calc（电子表格）类型。")


def _get_sheet(project: Dict[str, Any], sheet: int = 0) -> Dict[str, Any]:
    """按索引获取工作表。"""
    _ensure_calc(project)
    sheets = project.get("sheets", [])
    if sheet < 0 or sheet >= len(sheets):
        raise IndexError(f"工作表索引超出范围: {sheet}（共 {len(sheets)} 个）")
    return sheets[sheet]


def _validate_cell_ref(ref: str) -> str:
    """验证并规范化单元格引用（如 A1, B2, AA10）。"""
    match = re.match(r"^([A-Za-z]{1,3})(\d+)$", ref)
    if not match:
        raise ValueError(f"无效的单元格引用: {ref}。请使用如 A1、B2、AA10 的格式。")
    col = match.group(1).upper()
    row = int(match.group(2))
    return f"{col}{row}"


def add_sheet(
    project: Dict[str, Any],
    name: str = "Sheet",
    position: Optional[int] = None,
) -> Dict[str, Any]:
    """添加新工作表。"""
    _ensure_calc(project)
    if "sheets" not in project:
        project["sheets"] = []

    sheet = {"name": name, "cells": {}}
    sheets = project["sheets"]
    if position is not None and 0 <= position < len(sheets):
        sheets.insert(position, sheet)
    else:
        sheets.append(sheet)
    return sheet


def remove_sheet(project: Dict[str, Any], sheet: int) -> Dict[str, Any]:
    """按索引删除工作表。"""
    _ensure_calc(project)
    sheets = project.get("sheets", [])
    if sheet < 0 or sheet >= len(sheets):
        raise IndexError(f"工作表索引超出范围: {sheet}（共 {len(sheets)} 个）")
    removed = sheets.pop(sheet)
    return removed


def rename_sheet(project: Dict[str, Any], sheet: int, name: str) -> Dict[str, Any]:
    """重命名工作表。"""
    s = _get_sheet(project, sheet)
    s["name"] = name
    return s


def set_cell(
    project: Dict[str, Any],
    ref: str,
    value: Any,
    cell_type: str = "string",
    sheet: int = 0,
    formula: Optional[str] = None,
) -> Dict[str, Any]:
    """设置单元格的值。

    Args:
        project: 项目字典
        ref: 单元格引用（如 A1）
        value: 单元格值
        cell_type: 数据类型 —— string / float / boolean / formula
        sheet: 工作表索引
        formula: 公式字符串

    Returns:
        设置的单元格信息
    """
    ref = _validate_cell_ref(ref)
    s = _get_sheet(project, sheet)

    # 自动推断数据类型
    if cell_type == "string":
        try:
            float(value)
            cell_type = "float"
        except (ValueError, TypeError):
            pass

    cell = {"value": value, "type": cell_type}
    if formula:
        cell["formula"] = formula

    s["cells"][ref] = cell
    return {"ref": ref, "sheet": sheet, **cell}


def get_cell(
    project: Dict[str, Any],
    ref: str,
    sheet: int = 0,
) -> Dict[str, Any]:
    """获取单元格的值。"""
    ref = _validate_cell_ref(ref)
    s = _get_sheet(project, sheet)
    cell = s["cells"].get(ref)
    if cell is None:
        return {"ref": ref, "sheet": sheet, "value": None, "type": "empty"}
    return {"ref": ref, "sheet": sheet, **cell}


def clear_cell(
    project: Dict[str, Any],
    ref: str,
    sheet: int = 0,
) -> Dict[str, Any]:
    """清除单元格。"""
    ref = _validate_cell_ref(ref)
    s = _get_sheet(project, sheet)
    s["cells"].pop(ref, None)
    return {"ref": ref, "sheet": sheet, "cleared": True}


def set_range(
    project: Dict[str, Any],
    start_ref: str,
    data: List[List[Any]],
    sheet: int = 0,
) -> Dict[str, Any]:
    """批量写入一个矩形区域的数据。

    Args:
        project: 项目字典
        start_ref: 起始单元格引用（如 A1）
        data: 二维数据数组
        sheet: 工作表索引

    Returns:
        包含写入单元格数量的结果字典
    """
    start_ref = _validate_cell_ref(start_ref)
    s = _get_sheet(project, sheet)

    start_col = re.match(r"^([A-Za-z]+)", start_ref).group(1).upper()
    start_row = int(re.match(r"^[A-Za-z]+(\d+)$", start_ref).group(1))

    count = 0
    for ri, row in enumerate(data):
        for ci, value in enumerate(row):
            col = _num_to_col(_col_to_num(start_col) + ci)
            ref = f"{col}{start_row + ri}"
            cell_type = "string"
            try:
                float(value)
                cell_type = "float"
            except (ValueError, TypeError):
                pass
            s["cells"][ref] = {"value": value, "type": cell_type}
            count += 1

    return {"cells_set": count, "start": start_ref, "rows": len(data), "cols": max(len(r) for r in data) if data else 0}


def merge_cells(
    project: Dict[str, Any],
    start_ref: str,
    end_ref: str,
    sheet: int = 0,
) -> Dict[str, Any]:
    """标记合并单元格区域。"""
    start_ref = _validate_cell_ref(start_ref)
    end_ref = _validate_cell_ref(end_ref)
    s = _get_sheet(project, sheet)

    if "merged_cells" not in s:
        s["merged_cells"] = []

    merge_entry = {"start": start_ref, "end": end_ref}
    s["merged_cells"].append(merge_entry)
    return merge_entry


def list_sheets(project: Dict[str, Any]) -> List[Dict[str, Any]]:
    """列出所有工作表。"""
    _ensure_calc(project)
    sheets = project.get("sheets", [])
    result = []
    for i, sheet in enumerate(sheets):
        result.append({
            "index": i,
            "name": sheet.get("name", "unknown"),
            "cell_count": len(sheet.get("cells", {})),
        })
    return result


def get_sheet_data(project: Dict[str, Any], sheet: int = 0) -> Dict[str, Any]:
    """获取工作表的完整数据。"""
    return _get_sheet(project, sheet)


# ── 图表（Chart）支持 ───────────────────────────────────────────

CHART_TYPES = {
    "column": 51,       # xlColumnClustered 柱状图
    "bar": 57,          # xlBarClustered 横向条形图
    "line": 4,          # xlLine 折线图
    "pie": 5,           # xlPie 饼图
    "area": -4098,      # xlArea 面积图
    "scatter": -4169,   # xlXYScatter 散点图
    "doughnut": -4120,  # xlDoughnut 环形图
}

VALID_CHART_TYPES = tuple(CHART_TYPES.keys())


def add_chart(
    project: Dict[str, Any],
    chart_type: str = "column",
    data_range: str = "A1:A10",
    title: Optional[str] = None,
    legend: bool = True,
    sheet: int = 0,
) -> Dict[str, Any]:
    """向工作表添加图表定义。

    Args:
        project: 项目字典
        chart_type: 图表类型 —— column / bar / line / pie / area / scatter / doughnut
        data_range: 数据源单元格区域（如 "A1:D13"，含表头行）
        title: 图表标题
        legend: 是否显示图例
        sheet: 工作表索引

    Returns:
        图表定义字典
    """
    if chart_type not in CHART_TYPES:
        raise ValueError(
            f"不支持的图表类型: {chart_type}。可用: {', '.join(VALID_CHART_TYPES)}"
        )

    s = _get_sheet(project, sheet)
    if "charts" not in s:
        s["charts"] = []

    chart = {
        "type": chart_type,
        "data_range": data_range,
        "title": title or "",
        "legend": bool(legend),
    }
    s["charts"].append(chart)
    return {"sheet": sheet, **chart}


def list_charts(project: Dict[str, Any], sheet: int = 0) -> List[Dict[str, Any]]:
    """列出工作表上的所有图表。"""
    s = _get_sheet(project, sheet)
    return list(s.get("charts", []))


def add_image(
    project: Dict[str, Any],
    path: str,
    sheet: int = 0,
    left: float = 10,
    top: float = 10,
    width: float = 480,
    height: float = 270,
) -> Dict[str, Any]:
    """向工作表插入图片（如 matplotlib 生成的图表 PNG）。

    WPS 表格的 Chart 对象在后台模式下与单元格格式共存不稳定，
    推荐用 matplotlib 生成图表 PNG 后以图片形式插入（与 PPT 工作流一致）。

    Args:
        project: 项目字典
        path: 图片文件路径（png/jpg 等）
        sheet: 工作表索引
        left/top: 图片左上角位置（磅）
        width/height: 图片显示尺寸（磅）

    Returns:
        图片插入定义
    """
    if not path or not os.path.exists(path):
        raise FileNotFoundError(f"图片文件不存在: {path}")
    s = _get_sheet(project, sheet)
    if "images" not in s:
        s["images"] = []

    img = {
        "path": os.path.abspath(path),
        "left": float(left), "top": float(top),
        "width": float(width), "height": float(height),
    }
    s["images"].append(img)
    return {"sheet": sheet, **img}


# ── 区域格式（Format）支持 ─────────────────────────────────────

# 对齐方式（Excel 兼容常量）
ALIGN_MAP = {
    "left": -4131,      # xlLeft
    "center": -4108,    # xlCenter
    "right": -4152,     # xlRight
}


def format_range(
    project: Dict[str, Any],
    start_ref: str,
    end_ref: Optional[str] = None,
    bold: Optional[bool] = None,
    fill_color: Optional[str] = None,
    font_color: Optional[str] = None,
    border: bool = False,
    col_width: Optional[float] = None,
    number_format: Optional[str] = None,
    align: Optional[str] = None,
    sheet: int = 0,
) -> Dict[str, Any]:
    """对单元格区域应用格式。

    Args:
        project: 项目字典
        start_ref: 起始单元格（如 A1）
        end_ref: 结束单元格（如 F1）；为 None 时仅格式化起始单元格
        bold: 是否加粗
        fill_color: 填充色（十六进制 RRGGBB，如 "1A3C8B"）
        font_color: 字体色（十六进制 RRGGBB）
        border: 是否加细边框
        col_width: 列宽（字符单位）
        number_format: 数字格式（如 "#,##0.00"、"0.0%"、"¥#,##0"）
        align: 对齐方式 —— left / center / right
        sheet: 工作表索引

    Returns:
        格式定义字典
    """
    start_ref = _validate_cell_ref(start_ref)
    if end_ref:
        end_ref = _validate_cell_ref(end_ref)
    s = _get_sheet(project, sheet)

    if "formats" not in s:
        s["formats"] = []

    fmt = {
        "range": f"{start_ref}:{end_ref}" if end_ref else start_ref,
        "bold": bool(bold) if bold is not None else None,
        "fill_color": fill_color,
        "font_color": font_color,
        "border": bool(border),
        "col_width": float(col_width) if col_width is not None else None,
        "number_format": number_format,
        "align": align,
    }
    s["formats"].append(fmt)
    return {"sheet": sheet, **fmt}


def set_col_width(
    project: Dict[str, Any],
    col: str,
    width: float,
    sheet: int = 0,
) -> Dict[str, Any]:
    """设置整列列宽。

    Args:
        project: 项目字典
        col: 列字母（如 "A"）
        width: 列宽（字符单位）
        sheet: 工作表索引

    Returns:
        列宽设置结果
    """
    col = col.upper().strip()
    if not re.match(r"^[A-Za-z]{1,3}$", col):
        raise ValueError(f"无效的列引用: {col}")
    s = _get_sheet(project, sheet)

    if "col_widths" not in s:
        s["col_widths"] = {}
    s["col_widths"][col] = float(width)
    return {"sheet": sheet, "col": col, "width": float(width)}


# ── 汇总行（Summary）支持 ──────────────────────────────────────

SUMMARY_FUNCS = ("sum", "average", "count", "max", "min")
FORMULA_FUNCS = {
    "sum": "SUM",
    "average": "AVERAGE",
    "count": "COUNT",
    "max": "MAX",
    "min": "MIN",
}


def add_summary(
    project: Dict[str, Any],
    data_range: str,
    func: str = "sum",
    label: Optional[str] = "合计",
    label_ref: Optional[str] = None,
    sheet: int = 0,
) -> Dict[str, Any]:
    """在数据区域末尾自动生成汇总行（公式驱动）。

    解析 data_range（如 "B2:D13"），在最后一行下方为每一列写入公式，
    如 =SUM(B2:B13)。汇总行标签写入 label_ref（默认数据区域第一列下方）。

    Args:
        project: 项目字典
        data_range: 数据区域（如 "B2:D13"），列字母可多列、行号连续
        func: 汇总函数 —— sum / average / count / max / min
        label: 汇总行标签文本；传 None 时不写标签（适合追加其他列汇总）
        label_ref: 标签单元格（默认取数据区域第一列下方一格）
        sheet: 工作表索引

    Returns:
        写入的公式列表
    """
    if func not in FORMULA_FUNCS:
        raise ValueError(f"不支持的汇总函数: {func}。可用: {', '.join(SUMMARY_FUNCS)}")

    m = re.match(r"^([A-Za-z]{1,3})(\d+):([A-Za-z]{1,3})(\d+)$", data_range.strip())
    if not m:
        raise ValueError(
            f"无效的数据区域: {data_range}。请使用如 B2:D13 的连续区域格式。"
        )
    start_col, start_row, end_col, end_row = (
        m.group(1).upper(), int(m.group(2)),
        m.group(3).upper(), int(m.group(4)),
    )
    if end_row < start_row:
        raise ValueError(f"数据区域行号顺序错误: {data_range}")

    summary_row = end_row + 1
    func_name = FORMULA_FUNCS[func]

    # 标签单元格（默认数据区域第一列下方一格；label=None 时不设标签、
    # 不跳过任何公式单元格）
    if label_ref is None:
        label_ref = f"{start_col}{summary_row}" if label is not None else None
    if label_ref is not None:
        label_ref = _validate_cell_ref(label_ref)
    if label is not None:
        set_cell(project, label_ref, label, cell_type="string", sheet=sheet)

    # 为每一列写公式（跳过标签所在单元格）
    written = []
    col_i = _col_to_num(start_col)
    end_col_i = _col_to_num(end_col)
    while col_i <= end_col_i:
        col = _num_to_col(col_i)
        cell_ref = f"{col}{summary_row}"
        if cell_ref != label_ref:
            formula = f"={func_name}({col}{start_row}:{col}{end_row})"
            set_cell(project, cell_ref, 0.0, cell_type="float",
                     formula=formula, sheet=sheet)
            written.append({"ref": cell_ref, "formula": formula})
        col_i += 1

    return {"summary_row": summary_row, "func": func, "formulas": written}


# ── 辅助函数 ────────────────────────────────────────────────────

def _col_to_num(col: str) -> int:
    """列字母转数字（A=1, B=2, ..., Z=26, AA=27）。"""
    n = 0
    for c in col.upper():
        n = n * 26 + (ord(c) - ord("A") + 1)
    return n


def _num_to_col(n: int) -> str:
    """数字转列字母（1=A, 2=B, 26=Z, 27=AA）。"""
    result = ""
    while n > 0:
        n, m = divmod(n - 1, 26)
        result = chr(ord("A") + m) + result
    return result
