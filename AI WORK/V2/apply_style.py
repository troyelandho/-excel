#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Apply a strict visual-only V2 to an EXISTING workbook.

Never imports/runs either model generator. Never evaluates formulas. The input
file is read once; only styles.xml and whitelist attributes in sheet XML are
patched. All other ZIP members, including charts, relationships, shared strings,
calculation chain and properties, are copied byte for byte. Formula payloads and
caches are kept literally unchanged. A fail-closed verifier runs before delivery.
Python 3.10+; dependency: lxml. See README.md for use and limitations.
"""
from __future__ import annotations

import argparse
from collections import Counter
from copy import deepcopy
from datetime import datetime
import hashlib
import io
import json
import math
from pathlib import Path
import posixpath
import re
import sys
import unicodedata
import zipfile

from lxml import etree as ET

S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"s": S, "r": R}
PARSER = ET.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
HERE = Path(__file__).resolve().parent
DEFAULT_INPUT = HERE.parent / "V1" / "交付文件" / "产业园投资测算.xlsx"
OUTPUT_NAME = "产业园投资测算_V2_样式调整.xlsx"

# Only visual constants. There are no project defaults or model assumptions here.
THEME = {
    "font": "微软雅黑", "text": "334155", "title": "1F2937",
    "note": "64748B", "module": "4472C4", "header": "D9E2F3",
    "stripe": "F7F9FC", "total": "EEF2F6", "key": "EAF2FF",
    "input": "FFF9C4", "input_text": "1256A0", "border": "D9DEE7",
    "separator": "AAB4C5", "white": "FFFFFF",
}
TAB_COLORS = {"使用说明": "94A3B8", "核心参数输入": "4472C4",
              "指标与情景测算": "5B7FA6", "决策看板": "1F4E78"}
TOTAL_WORDS = ("合计", "小计", "总成本", "含预备费", "现金成本及经营税费", "累计所需资本金")
KEY_WORDS = ("项目总投资", "总可售面积", "总持有出租面积", "可经营面积",
             "项目投资税后现金流", "资本金现金流（完整出资理论）", "累计项目现金流",
             "累计项目折现现金流", "可用于偿债现金CFADS", "当年本息合计")
FONT_ALLOWED = {"name", "sz", "b", "color"}
ALIGN_ALLOWED = {"horizontal", "vertical", "wrapText"}
XF_VISUAL_ATTRS = {"fontId", "fillId", "borderId", "applyFont", "applyFill", "applyBorder", "applyAlignment"}


def q(name):
    return f"{{{S}}}{name}"


def local(node):
    return ET.QName(node).localname


def xml(data):
    return ET.fromstring(data, PARSER)


def sha(data):
    return hashlib.sha256(data).hexdigest()


def canon(node):
    """Namespace-aware tree equality, ignoring XML layout whitespace only."""
    if node is None:
        return None
    text = node.text
    if text is not None and not text.strip() and len(node):
        text = None
    return [node.tag, sorted(node.attrib.items()), text, [canon(c) for c in node]]


def package(data):
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        if len(z.namelist()) != len(set(z.namelist())):
            raise ValueError("ZIP 包中存在重复成员，拒绝处理")
        bad = z.testzip()
        if bad:
            raise ValueError(f"ZIP 校验失败：{bad}")
        return {i.filename: z.read(i) for i in z.infolist()}


def sheet_records(parts):
    rels = {n.get("Id"): n.get("Target") for n in xml(parts["xl/_rels/workbook.xml.rels"])}
    result = []
    for sh in xml(parts["xl/workbook.xml"]).find("s:sheets", NS):
        target = rels[sh.get(f"{{{R}}}id")]
        part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
        result.append({"name": sh.get("name"), "part": part, "state": sh.get("state", "visible")})
    return result


def shared_strings(parts):
    if "xl/sharedStrings.xml" not in parts:
        return []
    return ["".join(n.itertext()) for n in xml(parts["xl/sharedStrings.xml"])]


def cell_text(cell, strings):
    value = cell.find("s:v", NS)
    kind = cell.get("t")
    if kind == "s" and value is not None and value.text is not None:
        return strings[int(value.text)]
    if kind == "inlineStr":
        inline = cell.find("s:is", NS)
        return "".join(inline.itertext()) if inline is not None else ""
    if kind == "str" and value is not None:
        return value.text or ""
    return ""


def col_index(name):
    result = 0
    for c in name:
        result = result * 26 + ord(c) - 64
    return result


def cell_coord(address):
    match = re.fullmatch(r"([A-Z]+)([0-9]+)", address)
    if not match:
        raise ValueError(f"不支持的单元格地址：{address}")
    return col_index(match[1]), int(match[2])


def merge_bounds(reference):
    a, b = reference.split(":") if ":" in reference else (reference, reference)
    c1, r1 = cell_coord(a)
    c2, r2 = cell_coord(b)
    return c1, r1, c2, r2


def change_attr(tag, name, value):
    """Patch only one XML opening-tag attribute; leave other bytes untouched."""
    pattern = rb'(?<![A-Za-z0-9_:])' + name.encode() + rb'="[^"]*"'
    replacement = name.encode() + b'="' + str(value).encode() + b'"'
    if re.search(pattern, tag):
        return re.sub(pattern, lambda _: replacement, tag, count=1)
    suffix = 2 if tag.endswith(b"/>") else 1
    return tag[:-suffix] + b" " + replacement + tag[-suffix:]


def font_color(font):
    node = font.find("s:color", NS)
    return node.get("rgb", "")[-6:].upper() if node is not None else ""


def fill_color(fill):
    node = fill.find("s:patternFill/s:fgColor", NS)
    return node.get("rgb", "")[-6:].upper() if node is not None else ""


def set_child(parent, name, attributes):
    child = parent.find(f"s:{name}", NS)
    if child is None:
        child = ET.SubElement(parent, q(name))
    child.attrib.clear()
    child.attrib.update(attributes)
    return child


def font_projection(font):
    node = deepcopy(font)
    for child in list(node):
        if local(child) in FONT_ALLOWED:
            node.remove(child)
    return canon(node)


def fill_projection(fill):
    node = deepcopy(fill)
    for child in node.iter():
        if local(child) in {"fgColor", "bgColor", "color"}:
            child.attrib.clear()
        if local(child) == "patternFill" and child.get("patternType", "none") in {"none", "solid"}:
            child.attrib.pop("patternType", None)
    # Color children may be introduced to express a previously unfilled cell.
    for child in list(node.iter()):
        for sub in list(child):
            if local(sub) in {"fgColor", "bgColor", "color"}:
                child.remove(sub)
    return canon(node)


def border_projection(border):
    node = deepcopy(border)
    for edge in node:
        edge.attrib.pop("style", None)
        for child in list(edge):
            if local(child) == "color":
                edge.remove(child)
    return canon(node)


def xf_projection(xf):
    node = deepcopy(xf)
    for attr in XF_VISUAL_ATTRS:
        node.attrib.pop(attr, None)
    alignment = node.find("s:alignment", NS)
    if alignment is not None:
        for attr in ALIGN_ALLOWED:
            alignment.attrib.pop(attr, None)
        if not alignment.attrib and not len(alignment):
            node.remove(alignment)
    return canon(node)


def effective_invariants(style, style_id):
    xf = style.find("s:cellXfs", NS)[style_id]
    fonts, fills, borders = (style.find(f"s:{group}", NS) for group in ("fonts", "fills", "borders"))
    return {"xf": xf_projection(xf),
            "font_other": font_projection(fonts[int(xf.get("fontId", "0"))]),
            "fill_other": fill_projection(fills[int(xf.get("fillId", "0"))]),
            "border_other": border_projection(borders[int(xf.get("borderId", "0"))]),
            "numFmtId": xf.get("numFmtId", "0")}


class StyleRegistry:
    """Append minimal derivatives; retain every original style definition."""
    def __init__(self, raw):
        self.raw = raw
        self.root = xml(raw)
        self.groups = {n: list(self.root.find(f"s:{n}", NS)) for n in ("fonts", "fills", "borders", "cellXfs")}
        self.original_counts = {n: len(v) for n, v in self.groups.items()}
        self.caches = {n: {ET.tostring(x): i for i, x in enumerate(v)} for n, v in self.groups.items()}
        self.style_cache = {}

    def append(self, group, node):
        key = ET.tostring(node)
        if key not in self.caches[group]:
            self.caches[group][key] = len(self.groups[group])
            self.groups[group].append(node)
        return self.caches[group][key]

    def make(self, old_id, role, stripe=False, align="left", wrap=True):
        key = (old_id, role, stripe, align, wrap)
        if key in self.style_cache:
            return self.style_cache[key]
        xf = deepcopy(self.groups["cellXfs"][old_id])
        font = deepcopy(self.groups["fonts"][int(xf.get("fontId", "0"))])
        old_fill = self.groups["fills"][int(xf.get("fillId", "0"))]
        fill = deepcopy(old_fill)
        border = deepcopy(self.groups["borders"][int(xf.get("borderId", "0"))])
        color, size, bold = THEME["text"], 10.5, False
        background = THEME["stripe"] if stripe else THEME["white"]
        bottom, top = "thin", None
        if role == "title":
            color, size, bold, background, bottom = THEME["title"], 15, True, THEME["white"], None
        elif role == "module":
            color, size, bold, background = THEME["white"], 10.5, True, THEME["module"]
        elif role == "header":
            color, size, bold, background = THEME["title"], 10.5, True, THEME["header"]
        elif role == "note":
            color, size, bold, background, bottom = THEME["note"], 9.5, False, THEME["stripe"], None
        elif role == "input":
            color, background = THEME["input_text"], THEME["input"]
        elif role == "total":
            bold, background, top = True, THEME["total"], "medium"
        elif role == "key":
            bold, background, top = True, THEME["key"], "thin"
        elif role == "decision":
            size, bold, background, bottom = 14, True, THEME["key"], None
        elif role == "status":
            # Existing differential styles still supply dynamic green/red/gray.
            bold = True
        elif role == "label":
            bold = False

        # Only name/size/bold/color are changed. Italic, underline, scheme, family,
        # charset, strike, superscript etc. are inherited exactly from the source.
        set_child(font, "name", {"val": THEME["font"]})
        set_child(font, "sz", {"val": str(size)})
        set_child(font, "color", {"rgb": "FF" + color})
        if bold:
            set_child(font, "b", {})
        else:
            b = font.find("s:b", NS)
            if b is not None:
                font.remove(b)

        pattern = fill.find("s:patternFill", NS)
        if pattern is None or pattern.get("patternType", "none") not in {"none", "solid"}:
            raise ValueError("遇到渐变或特殊填充，不能在颜色白名单内安全处理")
        pattern.set("patternType", "solid")
        set_child(pattern, "fgColor", {"rgb": "FF" + background})
        # Retain any existing background attributes; add no non-color settings.

        # Keep edge existence and diagonal/outline flags. Only edge line and color
        # change. No full heavy grid: body cells have a light lower separator.
        for edge in border:
            name = local(edge)
            if name not in {"left", "right", "top", "bottom"}:
                continue
            line = top if name == "top" else bottom if name == "bottom" else None
            edge.attrib.pop("style", None)
            color_node = edge.find("s:color", NS)
            if color_node is not None:
                edge.remove(color_node)
            if line:
                edge.set("style", line)
                set_child(edge, "color", {"rgb": "FF" + (THEME["separator"] if line == "medium" else THEME["border"])})

        xf.set("fontId", str(self.append("fonts", font)))
        xf.set("fillId", str(self.append("fills", fill)))
        xf.set("borderId", str(self.append("borders", border)))
        xf.set("applyFont", "1")
        xf.set("applyFill", "1")
        xf.set("applyBorder", "1")
        xf.set("applyAlignment", "1")
        alignment = xf.find("s:alignment", NS)
        if alignment is None:
            # OOXML child order: alignment precedes protection and extension data.
            alignment = ET.Element(q("alignment"))
            pos = next((i for i, n in enumerate(xf) if local(n) in {"protection", "extLst"}), len(xf))
            xf.insert(pos, alignment)
        alignment.set("horizontal", align)
        alignment.set("vertical", "center")
        alignment.set("wrapText", "1" if wrap else "0")
        new_id = self.append("cellXfs", xf)
        self.style_cache[key] = new_id
        return new_id

    def fragment(self, node):
        result = ET.tostring(node, encoding="utf-8", with_tail=False)
        # New fragments inherit root bindings. Remove only redundant declarations.
        for prefix, uri in self.root.nsmap.items():
            key = b"xmlns" if prefix is None else b"xmlns:" + prefix.encode()
            result = result.replace(b" " + key + b'="' + uri.encode() + b'"', b"")
        return result

    def serialize(self):
        raw = self.raw
        for group, nodes in self.groups.items():
            added = nodes[self.original_counts[group]:]
            if not added:
                continue
            # Keep the original XML of EVERY existing definition literally intact.
            pattern = rb"<" + group.encode() + rb"\b[^>]*>.*?</" + group.encode() + rb">"
            matches = list(re.finditer(pattern, raw, re.S))
            if len(matches) != 1:
                raise ValueError(f"不支持的样式 XML 容器：{group}")
            match = matches[0]
            chunk = match.group()
            opening_end = chunk.index(b">") + 1
            opening = change_attr(chunk[:opening_end], "count", len(nodes))
            ending = b"</" + group.encode() + b">"
            replacement = opening + chunk[opening_end:-len(ending)] + b"".join(self.fragment(n) for n in added) + ending
            raw = raw[:match.start()] + replacement + raw[match.end():]
        return raw


def model_rows(sheet, strings):
    labels = {}
    for row in sheet.findall("s:sheetData/s:row", NS):
        for c in row:
            if c.get("r") == "A" + row.get("r"):
                labels[int(row.get("r"))] = cell_text(c, strings)
                break
    return labels


def classify(cell, sheet_name, labels, registry, strings):
    col, row = cell_coord(cell.get("r"))
    old = registry.groups["cellXfs"][int(cell.get("s", "0"))]
    font = registry.groups["fonts"][int(old.get("fontId", "0"))]
    fill = registry.groups["fills"][int(old.get("fillId", "0"))]
    text = cell_text(cell, strings)
    formula = cell.find("s:f", NS) is not None
    row_label = labels.get(row, "")
    size_node = font.find("s:sz", NS)
    size = float(size_node.get("val", "11")) if size_node is not None else 11
    original_alignment = old.find("s:alignment", NS)
    old_wrap = original_alignment is not None and original_alignment.get("wrapText") == "1"

    if row == 2 and col == 1 and not formula:
        return "title"
    if sheet_name == "决策看板":
        if cell.get("r") == "A5":
            return "decision"
        if row in {15, 16, 30, 31}:
            return "note"
        if 8 <= row <= 13 and (col in {2, 3, 7}):
            return "key"
        if 19 <= row <= 28 and col in {4, 6} or 35 <= row <= 64 and col == 4:
            return "status"
    if sheet_name == "核心参数输入" and fill_color(fill) in {"FFF2CC", THEME["input"]} and not formula:
        return "input"
    if sheet_name == "指标与情景测算":
        if 6 <= row <= 24 and col in {2, 4}:
            return "key"
        if 34 <= row <= 43 and col in {4, 6}:
            return "status"
        if row == 46 and col in {2, 4}:
            return "status"
        if col == 7 and row not in {4, 31, 33} or row in {25, 26, 27, 48, 49, 50} and col >= 2:
            return "note"
    if (font_color(font) == "FFFFFF" and fill_color(fill) == "203E56") or fill_color(fill) == THEME["header"]:
        return "header"
    if font_color(font) == "FFFFFF" and fill_color(fill) == THEME["module"]:
        return "module"
    if fill_color(fill) == "EAF0F5":
        return "total" if formula or col != 1 and text else "module"
    if any(w in row_label for w in KEY_WORDS) and not (sheet_name == "核心参数输入"):
        return "key"
    if any(w in row_label for w in TOTAL_WORDS) and not (sheet_name == "核心参数输入"):
        return "total"
    if old_wrap and (size <= 10 or font_color(font) == THEME["note"]) and (text or sheet_name in {"使用说明", "核心参数输入"}):
        return "note"
    if sheet_name == "核心参数输入" and col >= 4 and row < 80:
        return "note"
    return "label" if col == 1 and text else "body"


def estimated_units(text):
    return sum(2 if unicodedata.east_asian_width(c) in {"W", "F"} else 1 for c in text)


def patch_sheet(raw, sheet_name, registry, strings):
    sheet = xml(raw)
    cells = sheet.findall("s:sheetData/s:row/s:c", NS)
    labels = model_rows(sheet, strings)
    merges = [merge_bounds(n.get("ref")) for n in sheet.findall("s:mergeCells/s:mergeCell", NS)]
    roles = {c.get("r"): classify(c, sheet_name, labels, registry, strings) for c in cells}
    cell_map = {c.get("r"): c for c in cells}
    merge_at = {}
    for bounds in merges:
        c1, r1, c2, r2 = bounds
        anchor = next((c for c in cells if cell_coord(c.get("r")) == (c1, r1)), None)
        if anchor is not None:
            for c in cells:
                col, row = cell_coord(c.get("r"))
                if c1 <= col <= c2 and r1 <= row <= r2:
                    roles[c.get("r")] = roles[anchor.get("r")]
                    merge_at[c.get("r")] = bounds

    default_width = float(sheet.find("s:sheetFormatPr", NS).get("defaultColWidth", "8.43"))
    widths = {i: default_width for i in range(1, 16385)}
    for col in sheet.findall("s:cols/s:col", NS):
        for i in range(int(col.get("min")), int(col.get("max")) + 1):
            widths[i] = float(col.get("width", str(default_width)))
    heights = {int(r.get("r")): float(r.get("ht", "15")) for r in sheet.findall("s:sheetData/s:row", NS)}
    new_styles, new_heights, stats = {}, {}, Counter()
    for c in cells:
        address = c.get("r")
        col, row = cell_coord(address)
        role = roles[address]
        old_id = int(c.get("s", "0"))
        text = cell_text(c, strings)
        has_content = c.find("s:f", NS) is not None or c.find("s:v", NS) is not None or c.find("s:is", NS) is not None
        if not has_content and address not in merge_at and role not in {"module", "header"}:
            continue
        align = "center" if role == "header" else "left" if role in {"title", "module", "note", "label", "decision", "status"} or text and c.find("s:f", NS) is None else "right"
        if role == "input" and text:
            align = "left"
        wrap = role != "title"
        new_styles[address] = registry.make(old_id, role, stripe=row % 2 == 0, align=align, wrap=wrap)
        stats[role] += 1

        # Do not alter dashboard geometry: native chart anchors and pixel positions
        # remain exact because its widths/heights are also left unchanged.
        if sheet_name == "决策看板" or not text or role == "title":
            continue
        bounds = merge_at.get(address, (col, row, col, row))
        c1, r1, c2, r2 = bounds
        if (col, row) != (c1, r1):
            continue
        width = max(1.0, sum(widths[i] for i in range(c1, c2 + 1)) - 2)
        font_size = 9.5 if role == "note" else 10.5
        lines = sum(max(1, math.ceil(estimated_units(line) * (font_size / 10) / width)) for line in text.split("\n"))
        needed = font_size * 1.45 * lines + 5
        current = sum(heights.get(r, 15) for r in range(r1, r2 + 1))
        if needed > current + 0.5:
            # Change heights of existing rows only; never create, remove or move rows.
            addition = (needed - current) / (r2 - r1 + 1)
            for r in range(r1, r2 + 1):
                if r in heights:
                    new_heights[r] = max(new_heights.get(r, heights[r]), round(heights[r] + addition, 2))

    def cell_tag(match):
        tag = match.group()
        address = re.search(rb'\br="([A-Z]+[0-9]+)"', tag)
        if address and address.group(1).decode() in new_styles:
            return change_attr(tag, "s", new_styles[address.group(1).decode()])
        return tag

    result = re.sub(rb"<c\b[^>]*>", cell_tag, raw)

    def row_tag(match):
        tag = match.group()
        number = re.search(rb'\br="([0-9]+)"', tag)
        if number and int(number.group(1)) in new_heights:
            return change_attr(change_attr(tag, "ht", new_heights[int(number.group(1))]), "customHeight", "1")
        return tag

    result = re.sub(rb"<row\b[^>]*>", row_tag, result)
    zoom = 100 if sheet_name == "使用说明" else 90
    result = re.sub(rb"<sheetView\b[^>]*>", lambda m: change_attr(change_attr(m.group(), "showGridLines", "0"), "zoomScale", zoom), result)
    tab = TAB_COLORS.get(sheet_name, "7F9BC4")
    color_tag = b'<tabColor rgb="FF' + tab.encode() + b'"/>'
    if re.search(rb"<tabColor\b[^>]*/>", result):
        result = re.sub(rb"<tabColor\b[^>]*/>", lambda _: color_tag, result, count=1)
    elif re.search(rb"<sheetPr\b[^>]*>", result):
        match = re.search(rb"<sheetPr\b[^>]*>", result)
        result = result[:match.end()] + color_tag + result[match.end():]
    else:
        # tabColor is an explicitly permitted property. Add only its container.
        match = re.search(rb"<worksheet\b[^>]*>", result)
        result = result[:match.end()] + b"<sheetPr>" + color_tag + b"</sheetPr>" + result[match.end():]
    return result, {"roles": dict(stats), "styled_cells": len(new_styles), "row_heights_changed": len(new_heights), "column_widths_changed": 0}


def sheet_without_whitelist(sheet):
    node = deepcopy(sheet)
    for c in node.findall("s:sheetData/s:row/s:c", NS):
        c.attrib.pop("s", None)
    for row in node.findall("s:sheetData/s:row", NS):
        row.attrib.pop("ht", None)
        row.attrib.pop("customHeight", None)
    for col in node.findall("s:cols/s:col", NS):
        col.attrib.pop("width", None)
        col.attrib.pop("customWidth", None)
    for view in node.findall("s:sheetViews/s:sheetView", NS):
        view.attrib.pop("showGridLines", None)
        view.attrib.pop("zoomScale", None)
    sheet_pr = node.find("s:sheetPr", NS)
    if sheet_pr is not None:
        tab = sheet_pr.find("s:tabColor", NS)
        if tab is not None:
            sheet_pr.remove(tab)
        if not sheet_pr.attrib and not len(sheet_pr) and not sheet_pr.text:
            node.remove(sheet_pr)
    return canon(node)


def cells_raw(raw):
    result = {}
    for match in re.finditer(rb"<c\b[^>]*?(?:/>|>.*?</c>)", raw, re.S):
        payload = match.group()
        address = re.search(rb'\br="([A-Z]+[0-9]+)"', payload).group(1).decode()
        first_end = payload.index(b">") + 1
        opening = re.sub(rb'\s+s="[^"]*"', b"", payload[:first_end])
        result[address] = opening + payload[first_end:]
    return result


def nonempty(cell):
    return cell.find("s:f", NS) is not None or cell.find("s:v", NS) is not None or cell.find("s:is", NS) is not None


def make_baseline(parts, input_hash):
    records, sheets = sheet_records(parts), []
    for rec in records:
        root = xml(parts[rec["part"]])
        raw_cells = cells_raw(parts[rec["part"]])
        cells = []
        for c in root.findall("s:sheetData/s:row/s:c", NS):
            if nonempty(c):
                formula = c.find("s:f", NS)
                value = c.find("s:v", NS)
                cells.append({"address": c.get("r"), "type": c.get("t", "n"),
                              "formula": formula.text if formula is not None else None,
                              "formula_attributes": dict(formula.attrib) if formula is not None else None,
                              "cached_value": value.text if value is not None else None,
                              "payload_without_style": raw_cells[c.get("r")].decode("utf-8")})
        sheets.append({**rec, "cells": cells})
    return {"input_sha256": input_hash, "parts_sha256": {n: sha(d) for n, d in parts.items()}, "sheets": sheets}


def verify_packages(before, after):
    """Fail-closed verification over every member and every existing cell."""
    a, b = package(before), package(after)
    diffs = []
    checks = {}
    sheet_stats = []

    def record(name, count, detail=""):
        checks[name] = {"differences": count, "status": "通过" if count == 0 else "未通过", "detail": detail}
        if count:
            diffs.append({"check": name, "count": count, "detail": detail})

    same_members = set(a) == set(b)
    record("ZIP成员及顺序", int(not same_members) + int(list(a) != list(b)))
    if not same_members:
        return {"passed": False, "checks": checks, "differences": diffs}
    records_a, records_b = sheet_records(a), sheet_records(b)
    record("工作表名称数量顺序及可见状态", int(records_a != records_b))
    style_a, style_b = xml(a["xl/styles.xml"]), xml(b["xl/styles.xml"])
    cell_payload_diff = formulas_diff = caches_diff = counts_diff = number_diff = invariants_diff = sheets_diff = 0
    structure_diff = validation_diff = cf_diff = 0
    for rec in records_a:
        part = rec["part"]
        root_a, root_b = xml(a[part]), xml(b[part])
        cells_a = {c.get("r"): c for c in root_a.findall("s:sheetData/s:row/s:c", NS)}
        cells_b = {c.get("r"): c for c in root_b.findall("s:sheetData/s:row/s:c", NS)}
        raw_a, raw_b = cells_raw(a[part]), cells_raw(b[part])
        cell_payload_diff += sum(raw_a.get(addr) != raw_b.get(addr) for addr in set(raw_a) | set(raw_b))
        f_a = {addr: canon(c.find("s:f", NS)) for addr, c in cells_a.items() if c.find("s:f", NS) is not None}
        f_b = {addr: canon(c.find("s:f", NS)) for addr, c in cells_b.items() if c.find("s:f", NS) is not None}
        formulas_diff += sum(f_a.get(addr) != f_b.get(addr) for addr in set(f_a) | set(f_b))
        cache_a = {addr: [c.get("t"), canon(c.find("s:v", NS))] for addr, c in cells_a.items() if c.find("s:f", NS) is not None}
        cache_b = {addr: [c.get("t"), canon(c.find("s:v", NS))] for addr, c in cells_b.items() if c.find("s:f", NS) is not None}
        caches_diff += sum(cache_a.get(addr) != cache_b.get(addr) for addr in set(cache_a) | set(cache_b))
        counts_a = (len(cells_a), sum(nonempty(c) for c in cells_a.values()), len(f_a))
        counts_b = (len(cells_b), sum(nonempty(c) for c in cells_b.values()), len(f_b))
        counts_diff += int(counts_a != counts_b)
        for addr in set(cells_a) & set(cells_b):
            id_a, id_b = int(cells_a[addr].get("s", "0")), int(cells_b[addr].get("s", "0"))
            invariants_a, invariants_b = effective_invariants(style_a, id_a), effective_invariants(style_b, id_b)
            number_diff += int(invariants_a["numFmtId"] != invariants_b["numFmtId"])
            invariants_diff += int(invariants_a != invariants_b)
        sheets_diff += int(sheet_without_whitelist(root_a) != sheet_without_whitelist(root_b))
        for tag in ("dimension", "mergeCells", "sheetFormatPr", "sheetProtection", "protectedRanges", "autoFilter", "hyperlinks", "tableParts", "drawing", "legacyDrawing", "pageMargins", "pageSetup", "printOptions", "rowBreaks", "colBreaks", "headerFooter"):
            structure_diff += int(canon(root_a.find(f"s:{tag}", NS)) != canon(root_b.find(f"s:{tag}", NS)))
        for tag in ("dataValidations", "extLst"):
            validation_diff += int(canon(root_a.find(f"s:{tag}", NS)) != canon(root_b.find(f"s:{tag}", NS)))
        cf_diff += int([canon(n) for n in root_a.findall("s:conditionalFormatting", NS)] != [canon(n) for n in root_b.findall("s:conditionalFormatting", NS)])
        # Panes, selection, hidden/grouped rows/columns and all unlisted sheet flags
        # are already covered by the whole-sheet comparison above.
        sheet_stats.append({"name": rec["name"], "all_cell_nodes": counts_a[0], "nonempty_cells": counts_a[1], "formulas": counts_a[2],
                            "conditional_format_blocks": len(root_a.findall("s:conditionalFormatting", NS)),
                            "data_validation_rules": len(root_a.findall("s:dataValidations/s:dataValidation", NS))})

    record("逐单元格地址类型和值（含文字及错误标记）", cell_payload_diff, "单元格原始XML仅剔除样式索引后逐字节比较")
    record("公式字符串及公式属性", formulas_diff, "逐字符比较，同时比较共享/数组公式属性")
    record("原有公式缓存及缓存类型", caches_diff)
    record("内容单元格及公式数量", counts_diff)
    record("所有单元格有效数字格式", number_diff)
    record("有效样式中非白名单属性（含保护缩进旋转等）", invariants_diff)
    record("工作表完整非白名单XML", sheets_diff)
    record("合并隐藏冻结保护打印分页及对象关系", structure_diff)
    record("数据验证及扩展设置", validation_diff)
    record("条件格式公式阈值范围优先级停止规则", cf_diff)

    style_nonwhite_diff = 0
    reduced_b = deepcopy(style_b)
    for name in ("fonts", "fills", "borders", "cellXfs"):
        ga, gb = style_a.find(f"s:{name}", NS), reduced_b.find(f"s:{name}", NS)
        if ga is None or gb is None or len(gb) < len(ga):
            style_nonwhite_diff += 1
            continue
        style_nonwhite_diff += sum(canon(n) != canon(gb[i]) for i, n in enumerate(ga))
        for child in list(gb)[len(ga):]:
            gb.remove(child)
        gb.attrib.clear()
        gb.attrib.update(ga.attrib)
    style_nonwhite_diff += int(canon(style_a) != canon(reduced_b))
    record("原样式定义命名样式数字格式表及条件格式差异样式", style_nonwhite_diff)

    allowed_parts = {"xl/styles.xml"} | {r["part"] for r in records_a}
    other_changed = [n for n in a if n not in allowed_parts and a[n] != b[n]]
    record("其余ZIP成员逐字节保持", len(other_changed), ", ".join(other_changed))
    charts = [n for n in a if re.fullmatch(r"xl/charts/chart[0-9]+\.xml", n)]
    drawings = [n for n in a if n.startswith("xl/drawings/")]
    record("图表类型标题系列引用位置及全部设置", sum(a[n] != b[n] for n in charts + drawings), "图表及绘图XML、关系文件逐字节比较")
    record("命名范围计算设置及工作簿保护", int(a["xl/workbook.xml"] != b["xl/workbook.xml"]))
    record("超链接批注表格外部关系及其他对象", len(other_changed), "相关对象文件及全部关系文件均包含于逐字节检查")
    return {"passed": not diffs, "non_style_differences": sum(x["count"] for x in diffs),
            "checks": checks, "differences": diffs, "sheet_counts": sheet_stats,
            "totals": {"sheets": len(records_a), "all_cell_nodes": sum(s["all_cell_nodes"] for s in sheet_stats),
                       "nonempty_cells": sum(s["nonempty_cells"] for s in sheet_stats),
                       "formulas": sum(s["formulas"] for s in sheet_stats), "charts": len(charts),
                       "formula_caches": sum(s["formulas"] for s in sheet_stats), "zip_parts": len(a)},
            "byte_identical_parts": [n for n in a if a[n] == b[n]],
            "changed_parts": [n for n in a if a[n] != b[n]]}


def choose_output(directory, filename):
    target = directory / filename
    if not target.exists():
        return target
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    for serial in range(10000):
        suffix = f"_{stamp}" + (f"_{serial:02d}" if serial else "")
        target = directory / (Path(filename).stem + suffix + Path(filename).suffix)
        if not target.exists():
            return target
    raise FileExistsError("无法分配唯一输出名称")


def write_new(path, data):
    with path.open("xb") as f:
        f.write(data)


def json_bytes(data):
    return json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")


def apply_style(source, output_dir, filename=OUTPUT_NAME):
    source = Path(source).resolve(strict=True)
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    output = choose_output(output_dir, filename)
    if source == output.resolve():
        raise ValueError("输出不能是输入文件")
    original = source.read_bytes()
    original_hash = sha(original)
    parts = package(original)
    # Reject macro/digital-signature packages rather than discard their features.
    if any(n.startswith("_xmlsignatures/") or n.endswith("vbaProject.bin") for n in parts):
        raise ValueError("本纯样式脚本仅处理未签名的普通xlsx，不处理宏或签名文件")
    baseline = make_baseline(parts, original_hash)
    process_dir = output_dir / "_过程文件"
    process_dir.mkdir(exist_ok=True)
    baseline_file = choose_output(process_dir, "内容基准_" + output.stem + ".json")
    write_new(baseline_file, json_bytes(baseline))

    registry = StyleRegistry(parts["xl/styles.xml"])
    strings = shared_strings(parts)
    changed, summary = {}, {}
    for rec in sheet_records(parts):
        changed[rec["part"]], summary[rec["name"]] = patch_sheet(parts[rec["part"]], rec["name"], registry, strings)
    changed["xl/styles.xml"] = registry.serialize()
    stream = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(original)) as zin, zipfile.ZipFile(stream, "w") as zout:
        zout.comment = zin.comment
        for item in zin.infolist():
            zout.writestr(deepcopy(item), changed.get(item.filename, parts[item.filename]))
    candidate = stream.getvalue()
    verification = verify_packages(original, candidate)
    verification.update({"input": str(source), "output": str(output), "baseline_file": str(baseline_file),
                         "input_sha256_before": original_hash, "input_sha256_after": sha(source.read_bytes()),
                         "output_sha256": sha(candidate), "visual_changes": summary,
                         "processing": "OOXML白名单样式补丁；不运行模型生成代码；不计算或改写公式", "recalculated": False})
    original_intact = verification["input_sha256_after"] == original_hash
    verification["checks"]["原V1文件SHA256"] = {"differences": int(not original_intact), "status": "通过" if original_intact else "未通过"}
    if not original_intact:
        verification["passed"] = False
        verification["non_style_differences"] += 1
        verification["differences"].append({"check": "原文件在执行期间被其他程序修改", "count": 1})
    report = choose_output(output_dir, output.stem + "_验证结果.json")
    write_new(report, json_bytes(verification))
    if not verification["passed"]:
        raise RuntimeError(f"非样式差异不为0，未输出工作簿。详见：{report}")
    write_new(output, candidate)
    # Read the actual saved file and repeat the preservation checks.
    saved = output.read_bytes()
    if saved != candidate or not verify_packages(original, saved)["passed"] or sha(source.read_bytes()) != original_hash:
        raise RuntimeError("保存后验证失败，请勿使用输出文件")
    print(json.dumps({"output": str(output), "verification": str(report), "passed": True,
                      "non_style_differences": 0, "totals": verification["totals"]}, ensure_ascii=False))
    return output, report


def main():
    parser = argparse.ArgumentParser(description="对现有产业园工作簿只调整白名单视觉样式，保留公式缓存及全部模型内容")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT, help="当前已保存的原Excel，不读取默认参数")
    parser.add_argument("--output-dir", type=Path, default=HERE)
    parser.add_argument("--filename", default=OUTPUT_NAME)
    parser.add_argument("--verify-only", type=Path, metavar="V2_FILE", help="只比较原文件与指定新版，不生成或修改Excel")
    args = parser.parse_args()
    if Path(args.filename).name != args.filename or not args.filename.lower().endswith(".xlsx"):
        parser.error("--filename 必须是普通xlsx文件名")
    if args.verify_only:
        result = verify_packages(args.input.read_bytes(), args.verify_only.read_bytes())
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["passed"] else 1
    apply_style(args.input, args.output_dir, args.filename)
    return 0


if __name__ == "__main__":
    sys.exit(main())
