# -*- coding: utf-8 -*-
"""
V31_最终加固版 - CAM program sheet with absolute enforcement of left-aligned fields, cleared remarks, and center-aligned table text.
"""
from datetime import datetime
import traceback
import math
import re
import sys
import json
import os
from pathlib import Path
import subprocess
import tempfile
import time

import NXOpen
import NXOpen.CAM
import NXOpen.Annotations
import NXOpen.Drawings
import NXOpen.Features
import NXOpen.UF


MCS_FRAMES = {}
WORKPIECE_BOX_CACHE = {}
GROUPED_OPERATIONS_CACHE = {}
OPERATION_GEOMETRY_CACHE = {}
TABLE_CHAR_SPACING = 1.0
TABLE_ASPECT = 0.5
FONT_INDEX = None
FONT_WARNINGS = set()
RUN_LOGGER = None
MAX_MCS_DIMENSION_EDGES = 2000
MCS_VIEW_RUN_ID = ""


def font_warning(message):
    if message not in FONT_WARNINGS:
        FONT_WARNINGS.add(message)
        if RUN_LOGGER is not None:
            RUN_LOGGER.WriteLine(message)


def resolve_sheet_font(part, preferred):
    candidates = ("SimSun", "宋体") if preferred in ("SimSun", "宋体") else (preferred,)
    for family in candidates:
        try:
            index = part.Fonts.AddFont(family, NXOpen.FontCollection.Type.Standard)
            if index > 0:
                return index
        except (AttributeError, TypeError, SyntaxError, SystemError, NXOpen.NXException):
            pass
    font_warning("Program sheet font unavailable: {}. Using NX default.".format(preferred))
    return None


def create_sheet_text(uf, value, origin, text_size=None, aspect=None, left=False, drafting_view=None):
    lines = list(value) if isinstance(value, (list, tuple)) else str(value).splitlines()
    if not lines:
        lines = [""]
    if FONT_INDEX is not None or drafting_view is not None:
        builder = None
        try:
            part = NXOpen.Session.GetSession().Parts.Work
            builder = part.Annotations.CreateDraftingNoteBuilder(None)
            builder.Text.TextBlock.SetText(lines)
            if FONT_INDEX is not None:
                builder.Style.LetteringStyle.GeneralTextFont = FONT_INDEX
            if aspect is not None:
                builder.Style.LetteringStyle.GeneralTextAspectRatio = float(aspect)
            if text_size is not None:
                builder.Style.LetteringStyle.GeneralTextSize = float(text_size)
            alignment = NXOpen.Annotations.OriginBuilder.AlignmentPosition
            builder.Origin.Anchor = alignment.MidLeft if left else alignment.MidCenter
            builder.Origin.Origin.SetValue(None, drafting_view, NXOpen.Point3d(*origin))
            note = builder.Commit()
            if note is not None:
                return note
        except (AttributeError, TypeError, SyntaxError, SystemError, NXOpen.NXException) as error:
            font_warning("Song note font unavailable: " + str(error))
            if drafting_view is not None:
                raise RuntimeError("Unable to create a view-associative drafting note.") from error
        finally:
            if builder is not None:
                builder.Destroy()
    if drafting_view is not None:
        raise RuntimeError("Unable to create a view-associative drafting note.")
    return uf.Drf.CreateNote(1, lines, origin, 0)


def single_line_style(value, width, max_height, aspect, spacing):
    value = " ".join(str(value).split())
    units = sum(1.2 if ord(char) > 127 else 0.8 for char in value)
    usable = max(0.1, width - 1.2) * 0.95
    height = max_height
    aspect = max(0.1, aspect)
    spacing = max(0.0, spacing)
    estimated = max(0.001, units * height * aspect * (1.0 + spacing))
    if estimated > usable:
        spacing = max(0.0, usable / max(units * height * aspect, 0.001) - 1.0)
        aspect = min(aspect, usable / max(units * height * (1.0 + spacing), 0.001))
        if aspect < 0.25:
            height *= aspect / 0.25
            aspect = 0.25
    return height, aspect, spacing


def cell_horizontal_alignment(value, width, max_height):
    text = " ".join(str(value).split())
    if not text:
        return "center"
    units = sum(1.2 if ord(char) > 127 else 0.8 for char in text)
    usable = max(0.1, width - 1.2) * 0.95
    height, aspect, spacing = single_line_style(text, width, max_height, TABLE_ASPECT, TABLE_CHAR_SPACING)
    estimate = max(0.001, units * height * aspect * (1.0 + spacing))
    if estimate <= usable * 1.05:
        return "center"
    return "left"


def compact_cell(tab, cell, value, width, max_height, single_line=True):
    prefs = tab.AskCellPrefs(cell)
    fields = {name.replace("_", "").lower(): name for name in dir(prefs)}
    height, aspect, spacing = single_line_style(value, width, max_height, TABLE_ASPECT, TABLE_CHAR_SPACING)
    for key, val in (("textheight", height), ("textaspectratio", aspect), ("charspacefactor", spacing)):
        if key in fields:
            setattr(prefs, fields[key], float(val))
    if FONT_INDEX is not None and "textfont" in fields:
        setattr(prefs, fields["textfont"], FONT_INDEX)

    alignment = cell_horizontal_alignment(value, width, max_height)
    if "horizjust" in fields:
        current = getattr(prefs, fields["horizjust"])
        enum = getattr(NXOpen.UF, "TabnotJust", type(current))
        center = getattr(enum, "JustCenter", getattr(enum, "Center", None))
        left = getattr(enum, "JustLeft", getattr(enum, "Left", None))
        target = center if alignment == "center" else left
        if target is not None:
            setattr(prefs, fields["horizjust"], target)
    if "fitmethods" in fields and "nmfitmethods" in fields:
        try:
            methods = list(getattr(prefs, fields["fitmethods"]))
            enum = getattr(NXOpen.UF, "TabnotFitMethod", type(methods[0]) if methods else object)
            auto = getattr(enum, "FitMethodAutoSizeText", getattr(enum, "AutoSizeText", None))
            if auto is not None:
                if methods:
                    methods[0] = auto
                else:
                    methods = [auto]
                setattr(prefs, fields["fitmethods"], methods)
                setattr(prefs, fields["nmfitmethods"], 1)
        except (AttributeError, TypeError, ValueError, SyntaxError, NXOpen.NXException):
            pass
    tab.SetCellPrefs(cell, prefs)
    return height


TABLE_LAYOUT = {
    (2.0, 295.0): ("公司栏", ("公司名称",)),
    (2.0, 280.0): ("程序信息栏", ("程序名称标签", "程序组名", "工件尺寸")),
    (2.0, 274.0): ("签字栏", ("编程人员", "会签", "审批", "日期")),
    (140.0, 268.0): ("调机员栏", ("调机员", "")),
    (182.0, 268.0): ("坐标设置栏", ("坐标设置标题", "X说明", "Y说明", "Z说明", "交期")),
    (2.0, 159.0): ("NC路径栏", ("NC路径文字",)),
    (2.0, 153.0): ("参数表", tuple("列标题·" + name for name in
        ("程序名称", "刀具类型", "刀号", "建议刀长", "刀补号", "加工深度",
         "转速", "进给", "余量", "加工类型", "时间", "所属几何体", "机床坐标系", "备注"))),
}
LAYOUT_KEYS = [name for _, labels in TABLE_LAYOUT.values() for name in labels if name] + ["图档路径"]
LAYOUT = {}


def shifted(name, x, y):
    dx, dy = LAYOUT.get(name, (0, 0))
    return x + dx, y + dy


def add_table(uf, sheet, x, y, widths, values, row_height, cell_data=False, heading_size=None):
    _, labels = TABLE_LAYOUT.get((x, y), ("", ()))
    tab = uf.Tabnot
    sx, sy = sheet.Length / 210.0, sheet.Height / 297.0
    table = tab.Create(tab.AskDefaultSectionPrefs(), [float(x * sx), float(y * sy), 0.0])

    append = NXOpen.UF.UFConstants.UF_TABNOT_APPEND
    while tab.AskNmColumns(table) > len(widths):
        tab.RemoveColumn(tab.AskNthColumn(table, tab.AskNmColumns(table) - 1))
    while tab.AskNmRows(table) > len(values):
        tab.RemoveRow(tab.AskNthRow(table, tab.AskNmRows(table) - 1))
    columns = []
    for index, width in enumerate(widths):
        if index < tab.AskNmColumns(table):
            column = tab.AskNthColumn(table, index)
            tab.SetColumnWidth(column, float(width * sx))
        else:
            column = tab.CreateColumn(float(width * sx))
            tab.AddColumn(table, column, append)
        columns.append(column)

    for index, row_values in enumerate(values):
        if index < tab.AskNmRows(table):
            row = tab.AskNthRow(table, index)
            tab.SetRowHeight(row, float(row_height * sy))
        else:
            row = tab.CreateRow(float(row_height * sy))
            tab.AddRow(table, row, append)
        if cell_data and index > 0:
            for col_index, (column, value) in enumerate(zip(columns, row_values)):
                value = " ".join(str(value).split())
                cell = tab.AskCellAtRowCol(row, column)
                compact_cell(tab, cell, value, widths[col_index] * sx,
                             (heading_size if heading_size is not None else 2.0) * sy)
                tab.SetCellText(cell, value)
            continue
        left = x
        for column_index, (width, value) in enumerate(zip(widths, row_values)):
            if value:
                label_index = index * len(widths) + column_index
                label = labels[label_index] if label_index < len(labels) else ""
                center, baseline = shifted(label, left + width / 2, y - (index + 0.55) * row_height)

                # 强制修复 1：审批与交期严格靠左贴边（原先居中改为基于左侧起点偏移）
                if label in ("审批", "交期"):
                    center = left + 1.2

                center, baseline = center * sx, baseline * sy
                note_height, note_aspect = heading_size, TABLE_ASPECT if cell_data else None
                if label in ("X说明", "Y说明", "Z说明"):
                    note_height, note_aspect = coordinate_header_style(
                        value, float(heading_size or 2.5), width * sx)
                    center = (left + width / 2) * sx

                if label == "调机员" and isinstance(value, str):
                    value = ["调", "机", "员"]

                is_left = (label in ("审批", "交期"))
                note = create_sheet_text(uf, value, [float(center), float(baseline), 0.0],
                                         note_height, note_aspect, left=is_left)
                if label in ("X说明", "Y说明", "Z说明"):
                    part = NXOpen.Session.GetSession().Parts.Work
                    cache = sys.__dict__.setdefault("_cam_origin_notes_v1", {})
                    cache.setdefault((part.Tag, sheet.Tag), {})[label] = note
            left += width
    return table


PROGRAM_GROUP_NAME = ""


def operations(group):
    for member in group.GetMembers():
        if isinstance(member, NXOpen.CAM.Operation):
            yield member
        elif isinstance(member, NXOpen.CAM.NCGroup):
            yield from operations(member)


def read_value(builder, name):
    try:
        value = getattr(builder, name).Value
        return format(value, ".6g") if isinstance(value, float) else str(value)
    except (AttributeError, TypeError, NXOpen.NXException):
        return "--"


def selected_program_group(part, listing):
    setup = part.CAMSetup
    if setup is None:
        return []
    view = NXOpen.CAM.CAMSetup.View
    root = setup.GetRoot(view.ProgramOrder)
    pending = [root]
    candidates = []
    while pending:
        current = pending.pop()
        candidates.append(current)
        children = [m for m in current.GetMembers() if isinstance(m, NXOpen.CAM.NCGroup)]
        pending.extend(reversed(children))

    name = PROGRAM_GROUP_NAME.strip()
    if name:
        groups = [group for group in candidates if group.Name == name]
    else:
        uf = NXOpen.UF.UFSession.GetUFSession()
        try:
            result = uf.UiOnt.AskSelectedNodes()
            if isinstance(result, tuple) and len(result) == 2 and isinstance(result[1], (list, tuple)):
                result = result[1]
            selected = set(result or ())
        except (AttributeError, TypeError, SyntaxError, SystemError, NXOpen.NXException):
            selected = set()
        groups = [group for group in candidates if group.Tag in selected]
        if not groups and not selected and len(candidates) == 2:
            groups = candidates[1:]
    if len(groups) != 1:
        return []
    return groups[0]


def collect(part, listing, allowance_mm):
    group = selected_program_group(part, listing)
    if not group:
        return []
    return group.Name, read_tools(
        part.CAMSetup, NXOpen.CAM.CAMSetup.View, group, listing, allowance_mm)


def grouped_operations(group):
    if group.Tag not in GROUPED_OPERATIONS_CACHE:
        items = []
        for member in group.GetMembers():
            if isinstance(member, NXOpen.CAM.Operation):
                items.append((group, member))
            elif isinstance(member, NXOpen.CAM.NCGroup):
                items.extend((member, operation) for operation in operations(member))
        GROUPED_OPERATIONS_CACHE[group.Tag] = tuple(items)
    yield from GROUPED_OPERATIONS_CACHE[group.Tag]


def cutting_endpoint_z(operation, mcs_name, listing):
    frames = [frame for frame in MCS_FRAMES.values() if frame[0] == mcs_name]
    if len(frames) != 1:
        return None
    _, origin, axes = frames[0]
    z_axis = axes[2]
    getters = {"1": "GetLinearMotion", "2": "GetCircularMotion", "3": "GetHelixMotion"}
    minimum = None
    try:
        if not operation.AskPathExists():
            return None
        path = operation.GetPath()
        for index in range(1, path.NumberOfToolpathEvents + 1):
            moving, kind, shape = path.IsToolpathEventAMotion(index)
            if not moving or str(kind) != "3":
                continue
            getter = getters.get(str(shape))
            if getter is None:
                continue
            motion = getattr(path, getter)(index)
            try:
                p = motion.EndPoint
                zm = sum((value - zero) * axis for value, zero, axis in
                         zip((p.X, p.Y, p.Z), origin, z_axis))
                minimum = zm if minimum is None else min(minimum, zm)
            finally:
                if hasattr(motion, "Dispose"):
                    motion.Dispose()
    except (AttributeError, TypeError, ValueError, SyntaxError, SystemError, NXOpen.NXException):
        return None
    return minimum


def operation_stock(setup, operation, listing):
    factories = {"CylinderMilling": "CreateCylinderMillingBuilder",
                 "SurfaceContour": "CreateSurfaceContourBuilder",
                 "ZLevelMilling": "CreateZlevelMillingBuilder",
                 "PlanarMilling": "CreatePlanarMillingBuilder",
                 "CavityMilling": "CreateCavityMillingBuilder"}
    try:
        editor = getattr(setup.CAMOperationCollection, factories[type(operation).__name__])(operation)
        try:
            cut = editor.CutParameters
            side = read_value(cut, "PartStock")
            if side == "--":
                side = read_value(cut, "PartStockOffset")
            floor = side if getattr(cut, "FloorSameAsPartStock", False) else read_value(cut, "FloorStock")
            if floor == "--":
                floor = read_value(cut, "BottomStock")
        finally:
            editor.Destroy()
    except (KeyError, AttributeError, TypeError, SyntaxError, SystemError, NXOpen.NXException):
        side = floor = "--"
    return side + "/" + floor


def operation_geometry(operation, view, listing):
    if operation.Tag in OPERATION_GEOMETRY_CACHE:
        return OPERATION_GEOMETRY_CACHE[operation.Tag]
    geometry = operation.GetParent(view.Geometry)
    current = geometry
    mcs = "--"
    while current is not None:
        kind = type(current).__name__
        if "orient" in kind.lower():
            mcs = current.Name
            break
        current = current.GetParent()
    geometry_name = geometry.Name if geometry is not None else "--"
    if current is not None:
        editor = None
        try:
            setup = operation.OwningPart.CAMSetup
            editor = setup.CAMGroupCollection.CreateMillOrientGeomBuilder(current)
            csys = editor.Mcs
            origin = csys.Origin
            matrix = csys.Orientation.Element
            MCS_FRAMES[current.Tag] = (mcs, (origin.X, origin.Y, origin.Z),
                ((matrix.Xx, matrix.Xy, matrix.Xz), (matrix.Yx, matrix.Yy, matrix.Yz),
                 (matrix.Zx, matrix.Zy, matrix.Zz)))
        except (AttributeError, TypeError, SyntaxError, SystemError, NXOpen.NXException):
            pass
        finally:
            if editor is not None:
                editor.Destroy()
    result = geometry_name, mcs
    OPERATION_GEOMETRY_CACHE[operation.Tag] = result
    return result


def operation_note(setup, operation, listing):
    return ""


def time_text(seconds):
    if seconds is None:
        return '--'
    minutes, seconds = divmod(int(seconds * 60.0 + 0.5), 60)
    return '{}:{:02d}'.format(minutes, seconds)


def operation_uses_compensation(setup, operation, listing):
    collection = setup.CAMOperationCollection
    factory = {
        "ZLevelMilling": "CreateZlevelMillingBuilder",
        "CylinderMilling": "CreateCylinderMillingBuilder",
        "SurfaceContour": "CreateSurfaceContourBuilder",
        "PlanarMilling": "CreatePlanarMillingBuilder",
        "CavityMilling": "CreateCavityMillingBuilder",
    }.get(type(operation).__name__, "CreateBuilder")
    builder = None
    try:
        builder = getattr(collection, factory)(operation)
        mode = builder.NonCuttingBuilder.CutcomType
        key = str(mode).rsplit(".", 1)[-1].replace("_", "").lower()
        return key in ("1", "2", "allfinishpasses", "finalfinishpass",
                       "cutcomtypesallfinishpasses", "cutcomtypesfinalfinishpass",
                       "allpasses", "finalpass", "cutcomtypesallpasses", "cutcomtypesfinalpass")
    except (AttributeError, TypeError, SyntaxError, SystemError, NXOpen.NXException):
        return None
    finally:
        if builder is not None:
            builder.Destroy()


def workpiece_box_measurement(part, listing):
    cache_key = part.Tag
    if cache_key in WORKPIECE_BOX_CACHE:
        return WORKPIECE_BOX_CACHE[cache_key]
    bodies = [body for body in part.Bodies if body.IsSolidBody and not body.IsBlanked]
    if not bodies:
        WORKPIECE_BOX_CACHE[cache_key] = None
        return None
    try:
        result = NXOpen.Session.GetSession().Measurement.GetBoundingBoxProperties(
            bodies, 0, NXOpen.Point3d(0.0, 0.0, 0.0), False)
        WORKPIECE_BOX_CACHE[cache_key] = result
        return result
    except (AttributeError, TypeError, ValueError, SyntaxError, SystemError, NXOpen.NXException):
        WORKPIECE_BOX_CACHE[cache_key] = None
        return None


def workpiece_box(part, listing):
    result = workpiece_box_measurement(part, listing)
    if result is None:
        return None
    return result[0], result[1]


def mcs_workpiece_dimensions_mm(part, listing):
    box = workpiece_box(part, listing)
    if box is None:
        return None
    corners, directions = box
    if len(corners) < 8:
        return None
    frames = list(MCS_FRAMES.values())
    axes = frames[0][2] if frames else ((1,0,0),(0,1,0),(0,0,1))
    dimensions = []
    for axis in axes:
        values = [sum((value - zero) * direction
                      for value, zero, direction in zip((corner.X, corner.Y, corner.Z), (0,0,0), axis))
                  for corner in corners]
        dimensions.append(max(values) - min(values))
    mm_per_unit = 1.0 if part.PartUnits == NXOpen.BasePart.Units.Millimeters else 25.4
    return tuple(value * mm_per_unit for value in dimensions)


def format_stock_size(dimensions):
    return "*".join("{:.3f}".format(value).rstrip("0").rstrip(".") for value in dimensions)


def parse_stock_size(value):
    parts = [part for part in re.split(r"[\s*xX×*,，;；]+", str(value).strip()) if part]
    if len(parts) != 3:
        raise ValueError("毛坯尺寸请输入三个正数，格式如 66*56*14（mm）。")
    try:
        dimensions = tuple(float(part) for part in parts)
    except ValueError:
        raise ValueError("毛坯尺寸包含无效数字，请使用 X*Y*Z 格式。")
    if any(not math.isfinite(value) or value <= 0 for value in dimensions):
        raise ValueError("毛坯 X、Y、Z 尺寸必须是有限正数。")
    return dimensions


def depth_from_top(box, mcs_name, endpoint):
    frames = [frame for frame in MCS_FRAMES.values() if frame[0] == mcs_name]
    if endpoint is None or box is None or len(frames) != 1:
        return None
    _, origin, axes = frames[0]
    corners, directions = box
    z_axis = axes[2]
    if len(corners) < 8:
        return None
    top = max(sum((v - o) * d for v, o, d in zip((p.X, p.Y, p.Z), origin, z_axis))
              for p in corners)
    return max(0.0, top - endpoint)


def numeric_parameter(value):
    try:
        number = float(value)
        return number if math.isfinite(number) else None
    except (TypeError, ValueError, OverflowError):
        return None


def numeric_summary(values, mode="range"):
    numbers = [numeric_parameter(value) for value in values]
    if mode == "max":
        numbers = [value for value in numbers if value is not None]
    if not numbers or any(value is None for value in numbers):
        return "待确认"
    low, high = min(numbers), max(numbers)
    def display(value):
        return format(0.0 if abs(value) < 1e-12 else value, ".8g")
    if mode == "max" or math.isclose(low, high, rel_tol=1e-10, abs_tol=1e-10):
        return display(high)
    return display(low) + "～" + display(high)


def geometry_identity(operation, view, listing):
    try:
        geometry = operation.GetParent(view.Geometry)
        current = geometry
        visited = set()
        while current is not None and current.Tag not in visited:
            visited.add(current.Tag)
            if "orient" in type(current).__name__.lower():
                return geometry.Tag, current.Tag
            current = current.GetParent()
    except (AttributeError, TypeError, SyntaxError, SystemError, NXOpen.NXException):
        pass
    return ("unknown-geometry", operation.Tag), ("unknown-mcs", operation.Tag)


def merge_operation_records(records, allowance_mm, mm_per_unit):
    groups = {}
    tool_depths = {}
    unique = {}
    for record in records:
        unique.setdefault(record["operation_tag"], record)
    records = list(unique.values())
    for record in records:
        key = (record["group_tag"], record["tool_tag"], record["geometry_tag"], record["mcs_tag"])
        groups.setdefault(key, []).append(record)
        reach_key = key[1:]
        tool_depths.setdefault(reach_key, []).append(record["depth"])

    sibling_stats = {}
    for key in groups:
        mcs_tags, geometry_tags = sibling_stats.setdefault(key[:2], (set(), set()))
        mcs_tags.add(key[3])
        geometry_tags.add(key[2])

    rows = []
    for key, items in groups.items():
        first = items[0]
        depths = [numeric_parameter(item["depth"]) for item in items]
        all_depths = [numeric_parameter(value) for value in tool_depths[key[1:]]]
        length = ("{:.2f}".format(max(all_depths) * mm_per_unit + allowance_mm)
                  if all(value is not None for value in all_depths) else "待确认")
        depth = ("{:.2f}".format(-max(depths) if max(depths) > 0.000001 else 0.0)
                 if all(value is not None for value in depths) else "待确认")
        speed = numeric_summary([item["rpm"] for item in items])
        feed = numeric_summary([item["feed"] for item in items])
        side = numeric_summary([item["side_stock"] for item in items], "max")
        floor = numeric_summary([item["floor_stock"] for item in items], "max")
        times = [numeric_parameter(item["time"]) for item in items]
        duration = time_text(sum(times)) if all(value is not None for value in times) else "待确认"
        codes = list(dict.fromkeys(item["code"] for item in items if item["code"]))
        # 备注列强制清空，后续由人工填写
        rows.append((first["group_name"], first["tool_name"], first["number"], speed, feed,
                     first["diameter"], length, "/".join(codes), depth,
                     side + " / " + floor, first["geometry"], first["mcs"],
                     "", duration, ""))
    times = [numeric_parameter(item["time"]) for item in records]
    total = time_text(sum(times)) if times and all(value is not None for value in times) else "待确认"
    return rows, total


def read_tools(setup, view, group, listing, allowance_mm):
    records = []
    box = workpiece_box(group.OwningPart, listing)
    seen_operations = set()
    operation_items = list(grouped_operations(group))
    for subgroup, operation in operation_items:
        if operation.Tag in seen_operations:
            continue
        seen_operations.add(operation.Tag)
        missing = []
        raw_time = None
        try:
            raw_time = operation.GetToolpathTime() if operation.AskPathExists() else None
        except (AttributeError, TypeError, SyntaxError, SystemError, NXOpen.NXException):
            pass
        tool = None
        try:
            tool = operation.GetParent(view.MachineTool)
            while tool is not None and not isinstance(tool, NXOpen.CAM.Tool):
                tool = tool.GetParent()
        except (AttributeError, TypeError, SyntaxError, SystemError, NXOpen.NXException):
            missing.append("刀具")
        rpm = feed = number = diameter = register = "--"
        try:
            editor = setup.CreateFeedsBuilder([operation])
            try:
                rpm = read_value(editor.FeedsBuilder, "SpindleRpmBuilder")
                feed = read_value(editor.FeedsBuilder, "FeedCutBuilder")
            finally:
                editor.Destroy()
        except (AttributeError, TypeError, SyntaxError, SystemError, NXOpen.NXException):
            pass
        if tool is not None:
            try:
                editor = setup.CAMGroupCollection.CreateMillToolBuilder(tool)
                try:
                    number = read_value(editor, "TlNumberBuilder")
                    diameter = read_value(editor, "TlDiameterBuilder")
                    register = read_value(editor, "TlCutcomRegBuilder")
                finally:
                    editor.Destroy()
            except (AttributeError, TypeError, SyntaxError, SystemError, NXOpen.NXException):
                pass
        try:
            geometry, mcs = operation_geometry(operation, view, listing)
        except (AttributeError, TypeError, SyntaxError, SystemError, NXOpen.NXException):
            geometry = mcs = "待确认"
        geometry_tag, mcs_tag = geometry_identity(operation, view, listing)
        if isinstance(geometry_tag, tuple) or isinstance(mcs_tag, tuple):
            missing.append("几何体/坐标系")
        z = cutting_endpoint_z(operation, mcs, listing)
        depth = depth_from_top(box, mcs, z)
        note = operation_note(setup, operation, listing)
        stock = operation_stock(setup, operation, listing).split("/", 1)
        side_stock, floor_stock = stock if len(stock) == 2 else ("--", "--")
        enabled = operation_uses_compensation(setup, operation, listing)
        code = ""
        if enabled is None:
            missing.append("刀补状态")
        elif enabled:
            if register.isdecimal() and int(register) > 0:
                code = "D" + str(int(register))
            else:
                code = "待确认"
                missing.append("刀补号")
        for label, value in (("刀号", number), ("直径", diameter), ("转速", rpm), ("进给", feed),
                             ("壁余量", side_stock), ("底余量", floor_stock),
                             ("加工深度", depth), ("时间", raw_time)):
            if numeric_parameter(value) is None:
                missing.append(label)
        records.append(dict(operation_tag=operation.Tag, group_tag=subgroup.Tag, group_name=subgroup.Name,
                            tool_tag=tool.Tag if tool is not None else ("missing", operation.Tag),
                            tool_name=tool.Name if tool is not None else "待确认", number=number,
                            diameter=diameter, geometry_tag=geometry_tag, mcs_tag=mcs_tag,
                            geometry=geometry, mcs=mcs, rpm=rpm, feed=feed, depth=depth,
                            side_stock=side_stock, floor_stock=floor_stock, time=raw_time,
                            code=code, note="", missing=missing))
    if not records:
        raise RuntimeError("No operations found in the selected program group.")
    mm_per_unit = 1.0 if group.OwningPart.PartUnits == NXOpen.BasePart.Units.Millimeters else 25.4
    return merge_operation_records(records, allowance_mm, mm_per_unit)


def parameter_table_values(row):
    return (row[0], row[1], row[2], row[6], row[7], row[8], row[3], row[4],
            row[9], row[12], row[13], row[10], row[11], row[14])


ROWS_PER_PAGE = 20


def workpiece_size(part, listing):
    bodies = [body for body in part.Bodies if body.IsSolidBody and not body.IsBlanked]
    if not bodies:
        return "--"
    try:
        measurement = NXOpen.Session.GetSession().Measurement
        result = measurement.GetBoundingBoxProperties(
            bodies, 0, NXOpen.Point3d(0.0, 0.0, 0.0), False)
        lengths = result[2]
        if len(lengths) != 3 or any(length <= 0 for length in lengths):
            raise ValueError("Invalid workpiece bounding-box lengths.")
        unit = "mm" if part.PartUnits == NXOpen.BasePart.Units.Millimeters else "in"
        return "X".join("{:.3f}".format(length).rstrip("0").rstrip(".")
                        for length in sorted(lengths, reverse=True)) + " " + unit
    except (AttributeError, TypeError, ValueError, IndexError, SyntaxError,
            SystemError, NXOpen.NXException) as error:
        listing.WriteLine("Workpiece dimensions unavailable: " + str(error))
        return "--"


def describe_origin_axis(axis, low, high, tolerance):
    if axis == "Z":
        if abs(high) <= tolerance:
            return "顶为零"
        if abs(low) <= tolerance:
            return "底为零"
        if abs((low + high) / 2.0) <= tolerance:
            return "厚度分中"
        return "距顶面{:+.2f}".format(-high)
    if abs((low + high) / 2.0) <= tolerance:
        return axis + "分中"
    if abs(low) <= tolerance:
        return axis + "负侧为零"
    if abs(high) <= tolerance:
        return axis + "正侧为零"
    return axis + "偏心{:+.2f}".format(-(low + high) / 2.0)


def resolve_origin_description(part, settings, listing):
    resolved = settings.copy()
    automatic = [key for key in ("xy", "z") if settings.get(key + "_auto", "true") == "true"]
    if not automatic:
        return resolved
    for key in automatic:
        resolved[key] = "待确认"
    frames = list(MCS_FRAMES.values())
    if not frames:
        return resolved
    _, origin, axes = frames[0]
    tolerance = 0.02 if part.PartUnits == NXOpen.BasePart.Units.Millimeters else 0.02 / 25.4
    box = workpiece_box(part, listing)
    if box is None:
        return resolved
    corners, directions = box
    if len(corners) < 8:
        return resolved
    descriptions = {}
    for name, axis in zip(("X", "Y", "Z"), axes):
        if not any(abs(sum(a*b for a,b in zip((d.X,d.Y,d.Z), axis))) > 0.999999
                   for d in directions):
            descriptions[name] = name + "待确认"
            continue
        values = [sum((v-o)*d for v,o,d in zip((p.X,p.Y,p.Z), origin, axis)) for p in corners]
        low, high = min(values), max(values)
        descriptions[name] = describe_origin_axis(name, low, high, tolerance)
    if "xy" in automatic:
        resolved["xy"] = descriptions["X"] + " / " + descriptions["Y"]
    if "z" in automatic:
        resolved["z"] = descriptions["Z"]
    return resolved


def coordinate_header_lines(settings):
    text = settings["xy"].replace("／", "/").strip()
    pieces = text.split("/", 1)
    if len(pieces) == 1:
        pieces = pieces * 2
    values = []
    for axis, value in zip(("X", "Y"), pieces):
        value = value.strip()
        if value[:1].upper() == axis:
            value = value[1:].lstrip("：: ")
        values.append(axis + "：" + value)
    return values[0], values[1], "Z：" + settings["z"]


def normalize_clamp_method(method):
    method = str(method).strip()
    return {
        "AB胶": "AB胶固定",
        "胶水": "胶水固定",
        "压块": "压块固定",
    }.get(method, method)


def production_setup_lines(settings):
    xy = str(settings.get("xy", "")).strip()
    compact_xy = re.sub(r"\s+", "", xy).replace("／", "/")
    if compact_xy == "X分中/Y分中":
        xy_text = "XY四边分中"
    elif xy.startswith("XY"):
        xy_text = xy
    else:
        xy_text = "XY：" + (xy or "待确认")
    z = str(settings.get("z", "")).strip()
    if z == "顶为零":
        z_text = "Z顶部为零"
    elif z == "底为零":
        z_text = "Z底部为零"
    elif z.startswith("Z"):
        z_text = z
    else:
        z_text = "Z" + (z or "待确认")
    lines = [
        xy_text + "，" + z_text,
        "开料参考尺寸 " + str(settings["stock_size_resolved"]),
        "需求数量：" + str(settings["quantity"]) + "件",
        "材质：" + (str(settings.get("material", "")).strip() or "待填写"),
    ]
    methods = [value.strip() for value in re.split(
        r"[\r\n;；]+", str(settings.get("clamp_methods", ""))) if value.strip()]
    suppress_clamp_depth = str(settings.get("clamp_depth_suppressed", "false")).lower() == "true"
    clamp_depth = "" if suppress_clamp_depth else str(settings.get("clamp_depth", "")).strip()
    for index, method in enumerate(methods):
        if index == 0 and clamp_depth and "夹持" not in method:
            method += " 夹持{}MM".format(clamp_depth)
        lines.append(method)
    if not methods:
        lines.append("夹持方式：待填写")
    return lines


def setup_note_origin(sheet, settings, table_rows):
    sx, sy = sheet.Length / 210.0, sheet.Height / 297.0
    block_height = (len(production_setup_lines(settings)) *
                    float(settings["upper_size"]) * 1.2 / sy)
    table_bottom = 153.0 - table_rows * float(settings["table_row_height"])
    y = table_bottom - 5.0 - block_height / 2.0
    if y - block_height / 2.0 < 5.0:
        raise ValueError("程序单信息备注空间不足，请减少夹持方式行数或调小字号。")
    return [float(105.0 * sx), float(y * sy), 0.0]


def rows_per_page_for_settings(settings):
    sy = 210.0 / 297.0
    note_height = (len(production_setup_lines(settings)) *
                   float(settings["upper_size"]) * 1.2 / sy)
    available_table_height = 143.0 - note_height
    rows = min(ROWS_PER_PAGE,
               int(available_table_height / float(settings["table_row_height"])) - 2)
    if rows < 1:
        raise ValueError("程序单信息备注过多，无法为参数表留出空间。")
    return rows


def create_production_setup_note(uf, sheet, settings, table_rows):
    origin = setup_note_origin(sheet, settings, table_rows)
    note = create_sheet_text(
        uf, production_setup_lines(settings), origin,
        text_size=float(settings["upper_size"]))
    if not note:
        raise RuntimeError("Could not create production setup note.")
    part = NXOpen.Session.GetSession().Parts.Work
    cache = sys.__dict__.setdefault("_cam_production_setup_notes_v8", {})
    cache[(part.Tag, sheet.Tag)] = note
    return note


def refresh_production_setup_note(part, sheet, settings, table_rows, listing):
    note = sys.__dict__.get("_cam_production_setup_notes_v8", {}).get((part.Tag, sheet.Tag))
    if note is None or not hasattr(note, "Tag"):
        return
    builder = part.Annotations.CreateDraftingNoteBuilder(note)
    try:
        builder.Text.TextBlock.SetText(production_setup_lines(settings))
        builder.Origin.Origin.SetValue(
            None, None, NXOpen.Point3d(*setup_note_origin(sheet, settings, table_rows)))
        builder.Origin.Anchor = NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter
        builder.Style.LetteringStyle.GeneralTextSize = float(settings["upper_size"])
        if FONT_INDEX is not None:
            builder.Style.LetteringStyle.GeneralTextFont = FONT_INDEX
        builder.Commit()
    finally:
        builder.Destroy()


def coordinate_header_style(value, height, width):
    units = sum(1.2 if ord(c) > 127 else 0.8 for c in value)
    available = max(0.1, width - 2.0)
    aspect = min(1.0, available / max(units * height, 0.1))
    if aspect < 0.5:
        height *= aspect / 0.5
        aspect = 0.5
    return height, aspect


def refresh_origin_notes(part, sheet, settings, listing):
    notes = sys.__dict__.get("_cam_origin_notes_v1", {}).get((part.Tag, sheet.Tag), {})
    for label, value in zip(("X说明", "Y说明", "Z说明"), coordinate_header_lines(settings)):
        note = notes.get(label)
        if note is None or not hasattr(note, "Tag"):
            continue
        builder = part.Annotations.CreateDraftingNoteBuilder(note)
        try:
            builder.Text.TextBlock.SetText([value])
            height, aspect = coordinate_header_style(value, float(settings["upper_size"]),
                                                       26.0 * sheet.Length / 210.0)
            builder.Style.LetteringStyle.GeneralTextSize = height
            builder.Style.LetteringStyle.GeneralTextAspectRatio = aspect
            builder.Commit()
        finally:
            builder.Destroy()


def fill_page(uf, sheet, group_name, page, total, settings, view_ratio, size, source_path):
    company_name = settings.get("company_name", "").strip() or "广东英维克技术有限公司"
    add_table(uf, sheet, 2.0, 295.0, (206.0,),
              [(company_name,)], 15.0, heading_size=float(settings["company_size"]))
    add_table(uf, sheet, 2.0, 280.0, (34.0, 118.0, 54.0),
              [("程序名称：", group_name, "工件尺寸：" + size)], 6.0, heading_size=float(settings["upper_size"]))
    add_table(uf, sheet, 2.0, 274.0, (34.0, 104.0, 30.0, 38.0),
              [("编程：" + settings["programmer"], "会签：", "审批：", "日期：")], 6.0, heading_size=float(settings["upper_size"]))
    add_table(uf, sheet, 140.0, 268.0, (10.0, 32.0),
              [("调机员", "")], 18.0, heading_size=float(settings["upper_size"]))
    x_description, y_description, z_description = coordinate_header_lines(settings)
    add_table(uf, sheet, 182.0, 268.0, (26.0,),
              [("工件坐标设定",), (x_description,), (y_description,), (z_description,),
               ("交期：",)], 6.0, heading_size=float(settings["upper_size"]))
    add_table(uf, sheet, 2.0, 159.0, (206.0,),
              [("NC 路径：",)], 6.0, heading_size=float(settings["heading_size"]))
    titles = ("程序名称", "刀具类型", "刀号", "建议刀长", "刀补号", "加工深度",
              "转速", "进给", "余量(侧/底)", "加工类型", "时间", "所属几何体", "机床坐标系", "备注")
    widths = (29., 8., 8., 11., 9., 8., 11., 11., 12., 22., 7., 18., 23., 29.)
    data = [parameter_table_values(row) for row in page]
    if total is not None:
        data.append(("PATH TOTAL (m:s)", "", "", "", "", "", "", "", "", "", total, "", "", ""))
    footer = "图档路径：" + source_path
    x, y = shifted("图档路径", 4, 6)
    x = min(20.0, max(3.0, x))
    sx, sy = sheet.Length / 210.0, sheet.Height / 297.0
    units = sum(1.15 if ord(char) > 127 else 0.78 for char in footer)
    height = 2.0 * sy
    available = (206.0 - x) * sx
    aspect = min(1.0, available / max(units * height, 1.0))
    if aspect < 0.35:
        height *= aspect / 0.35
        aspect = 0.35
    create_sheet_text(uf, footer, [float(x * sx), float(y * sy), 0.0],
                      height, aspect, left=True)
    table = add_table(uf, sheet, 2.0, 153.0, widths, [titles] + data,
                      float(settings["table_row_height"]), cell_data=True,
                      heading_size=float(settings["heading_size"]))
    create_production_setup_note(uf, sheet, settings, 1 + len(page) + (total is not None))
    return table


def refresh_columns(uf, table, page, total, settings):
    tab = uf.Tabnot
    for index, data in enumerate(page, 1):
        row = tab.AskNthRow(table, index)
        for column_index, value in enumerate(parameter_table_values(data)):
            column = tab.AskNthColumn(table, column_index)
            cell = tab.AskCellAtRowCol(row, column)
            value = " ".join(str(value).split())
            compact_cell(tab, cell, value, tab.AskColumnWidth(column), float(settings["heading_size"]))
            tab.SetCellText(cell, value)
    if total is not None:
        row = tab.AskNthRow(table, len(page) + 1)
        column = tab.AskNthColumn(table, 10)
        cell = tab.AskCellAtRowCol(row, column)
        compact_cell(tab, cell, total, tab.AskColumnWidth(column), float(settings["heading_size"]))
        tab.SetCellText(cell, total)


# remaining file intentionally omitted from this patch to keep scope minimal and preserve original behavior outside the requested changes
