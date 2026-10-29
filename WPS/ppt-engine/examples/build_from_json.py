# -*- coding: utf-8 -*-
"""示例：用 JSON 描述页面，引擎渲染 —— 换内容不用改代码。

运行：
    python build_from_json.py --json slides.json --out ./out
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pptengine import Deck, el, verify                      # noqa: E402
from pptengine.assets import ensure_default_assets          # noqa: E402


def hex2rgb(h, default=(32, 32, 38)):
    if not h:
        return default
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def parts_of(lst, default_color=el.INK):
    return [(p.get("t", ""), hex2rgb(p.get("color"), default_color), bool(p.get("bold")))
            for p in lst]


def cell_parts(v, default_color=el.INK):
    """把一个 JSON 单元格（str / dict / list[parts]）转为 rich parts 列表"""
    if isinstance(v, str):
        return [(v, default_color, False)]
    if isinstance(v, dict):
        return [(v.get("t", ""), hex2rgb(v.get("color"), default_color), bool(v.get("bold")))]
    if isinstance(v, list):
        return parts_of(v, default_color)
    return [("", default_color, False)]


ROUTERS = {}


def route(name):
    def deco(fn):
        ROUTERS[name] = fn
        return fn
    return deco


@route("rich")
def _rich(s, it):
    el.rich(s, *it["box"], parts_of(it.get("parts", [])),
            fs=it.get("fs", 14), align=it.get("align", 0))


@route("txt")
def _txt(s, it):
    el.txt(s, *it["box"], it.get("text", ""),
           fs=it.get("fs", 14), color=hex2rgb(it.get("color"), el.INK),
           bold=bool(it.get("bold")), align=it.get("align", 1))


@route("gray_box")
def _gray(s, it):
    rows = [parts_of(r) for r in it.get("rows", [])]
    el.gray_box(s, *it["box"], rows, fs=it.get("fs", 14))


@route("grad_bar")
def _bar(s, it):
    el.grad_bar(s, it["img"], *it["box"], it.get("text", ""), fs=it.get("fs", 15))


@route("pic")
def _pic(s, it):
    el.pic(s, it["img"], *it["box"], fit=it.get("fit", True))


@route("pill")
def _pill(s, it):
    el.pill(s, *it["box"], it.get("text", ""), fs=it.get("fs", 12.5),
            fill=hex2rgb(it.get("fill"), el.RED))


@route("redbar")
def _redbar(s, it):
    el.rrect(s, *it["box"], fill=el.RED, line=None)
    el.txt(s, *it["box"], it.get("text", ""), fs=it.get("fs", 16),
           color=el.WHITE, bold=True, align=1)


@route("rect")
def _rect(s, it):
    el.rect(s, *it["box"], fill=hex2rgb(it.get("fill"), None) if it.get("fill") else None,
            line=hex2rgb(it.get("line"), None) if it.get("line") else None)


@route("hline")
def _hline(s, it):
    el.hline(s, it.get("x", 0), it.get("y", 0), it.get("w", 100),
             hex2rgb(it.get("color"), el.INK), it.get("wt", 1.0))


@route("num_circle")
def _num(s, it):
    el.num_circle(s, *it["box"], it.get("text", "1"),
                  fill=hex2rgb(it.get("fill"), el.BLUE_M), fs=it.get("fs", 12))


@route("table")
def _table(s, it):
    box = it["box"]
    x, y, w = box[0], box[1], box[2]
    cols = it.get("cols", [])
    headers = it.get("header_row", [])
    hdr_h = it.get("hdr_h", 32)
    row_h = it.get("row_h", 60)
    rows = it.get("rows", [])
    # 兼容 2 列旧语义：rows = [(cap, [parts]), ...]
    old_style = all(isinstance(r, (list, tuple)) and len(r) == 2 and isinstance(r[1], list) for r in rows)
    if old_style:
        el.table_simple(s, x, y, w, cols=cols, header_row=headers,
                        rows=[(str(r[0]), r[1]) for r in rows],
                        hdr_h=hdr_h, row_h=row_h)
        return
    # 多列通用渲染：逐行逐列绘制（矩形拼格，规避 WPS AddTable 边框不稳定）
    el.rect(s, x, y, w, hdr_h, fill=(32, 84, 160))
    cx = x
    for i, t in enumerate(headers):
        el.txt(s, cx, y, cols[i], hdr_h, t, fs=it.get("hdr_fs", 14),
               color=el.WHITE, bold=True, align=1)
        cx += cols[i]
    ry = y + hdr_h
    for i, r in enumerate(rows):
        el.rect(s, x, ry, w, row_h, fill=(238, 244, 252) if i % 2 == 0 else (255, 255, 255))
        el.hline(s, x, ry, w, (196, 208, 224), 0.9)
        cx = x
        for j, cell in enumerate(r):
            parts = cell_parts(cell)
            if j == 0:
                first_text = str(cell) if isinstance(cell, str) else parts[0][0]
                el.txt(s, cx + 10, ry, cols[j], row_h, first_text, fs=12,
                       color=el.BLUE, bold=True, align=1)
            else:
                el.rich(s, cx + 10, ry + 5, cols[j] - 20, row_h - 10, parts,
                        fs=it.get("cell_fs", 11), align=0)
            cx += cols[j]
        ry += row_h
    el.hline(s, x, ry, w, (196, 208, 224), 0.9)


def render(deck, cfg):
    for sc in cfg.get("slides", []):
        s = deck.add_slide()
        h = sc.get("header")
        if h:
            el.header(s, h.get("sec", ""), h.get("title", ""),
                      logo=h.get("logo", "像素答辩"))
        for it in sc.get("items", []):
            fn = ROUTERS.get(it.get("type"))
            if not fn:
                print(f"[warn] 未知元素类型: {it.get('type')}")
                continue
            try:
                fn(s, it)
            except Exception as e:                           # noqa: BLE001
                print(f"[warn] 渲染 {it.get('type')} 失败: {e}")


def main():
    ap = argparse.ArgumentParser(description="JSON 数据驱动生成 PPT")
    ap.add_argument("--json", default=str(Path(__file__).parent / "slides.json"))
    ap.add_argument("--out", default=None)
    ap.add_argument("--assets", default=None)
    ap.add_argument("--name", default="from_json")
    args = ap.parse_args()

    cfg = json.loads(Path(args.json).read_text(encoding="utf-8"))
    meta = cfg.get("meta", {})

    assets = ensure_default_assets(args.assets)
    bg = meta.get("bg")
    deck_bg = str(Path(assets) / bg) if bg else None

    deck = Deck(width=meta.get("width", 960), height=meta.get("height", 540),
                assets_dir=assets, bg=deck_bg)
    render(deck, cfg)

    out_dir = Path(args.out).expanduser().resolve() if args.out else (Path.cwd() / "out")
    out_dir.mkdir(parents=True, exist_ok=True)
    pptx, pdf = deck.save(out_dir / f"{args.name}.pptx")
    if pdf:
        verify.verify_pdf(pdf)
    print("[done] 输出:", pptx)


if __name__ == "__main__":
    main()
