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


def projected_hull(points):
    def half(sequence):
        chain = []
        for x, y in sequence:
            while len(chain) >= 2:
                a, b = chain[-2:]
                if (b[0] - a[0]) * (y - a[1]) > (b[1] - a[1]) * (x - a[0]):
                    break
                chain.pop()
            chain.append((x, y))
        return chain
    points = sorted(set(points))
    return half(points)[:-1] + half(reversed(points))[:-1]


def ray_exit(hull, x, y, dx, dy):
    hits = []
    for a, b in zip(hull, hull[1:] + hull[:1]):
        ux, uy = b[0] - a[0], b[1] - a[1]
        det = dx * uy - dy * ux
        if abs(det) < 1e-9:
            continue
        qx, qy = a[0] - x, a[1] - y
        distance = (qx * uy - qy * ux) / det
        fraction = (qx * dy - qy * dx) / det
        if distance >= 0 and -1e-8 <= fraction <= 1.0 + 1e-8:
            hits.append(distance)
    return max(hits) if hits else None


def mcs_view_unique_name(name):
    return '{}_{}'.format(name, MCS_VIEW_RUN_ID)


def mcs_view_xyz(point):
    return (float(point.X), float(point.Y), float(point.Z))


def mcs_view_add(a, b, factor=1.0):
    return tuple((x + factor * y for x, y in zip(a, b)))


def mcs_view_dot(a, b):
    return sum((x * y for x, y in zip(a, b)))


def mcs_view_unit(vector):
    length = math.sqrt(mcs_view_dot(vector, vector))
    if length <= 1e-12:
        raise ValueError('A CAM MCS axis has zero length.')
    return tuple((value / length for value in vector))


def mcs_view_cam_mcs(part, log):
    setup = part.CAMSetup
    if setup is None:
        raise RuntimeError('The work part has no CAM setup.')
    pending = [setup.GetRoot(NXOpen.CAM.CAMSetup.View.ProgramOrder)]
    seen = set()
    while pending:
        item = pending.pop()
        if item.Tag in seen:
            continue
        seen.add(item.Tag)
        if isinstance(item, NXOpen.CAM.NCGroup):
            pending.extend(reversed(list(item.GetMembers())))
            continue
        try:
            parent = item.GetParent(NXOpen.CAM.CAMSetup.View.Geometry)
        except (AttributeError, NXOpen.NXException):
            continue
        visited = set()
        while parent is not None and parent.Tag not in visited:
            visited.add(parent.Tag)
            if 'orient' in type(parent).__name__.lower():
                builder = setup.CAMGroupCollection.CreateMillOrientGeomBuilder(parent)
                try:
                    csys = builder.Mcs
                    matrix = csys.Orientation.Element
                    axes = ((matrix.Xx, matrix.Xy, matrix.Xz), (matrix.Yx, matrix.Yy, matrix.Yz), (matrix.Zx, matrix.Zy, matrix.Zz))
                    frame = (parent.Name, mcs_view_xyz(csys.Origin), axes)
                    return frame
                finally:
                    builder.Destroy()
            parent = parent.GetParent()
    raise RuntimeError('No CAM machining coordinate system was found.')


def mcs_view_choose_scale(part, session):
    bodies = [body for body in part.Bodies if body.IsSolidBody and (not body.IsBlanked)]
    if not bodies:
        raise RuntimeError('No visible solid bodies were found.')
    measurement = session.Measurement.GetBoundingBoxProperties(bodies, 0, NXOpen.Point3d(0.0, 0.0, 0.0), False)
    lengths = [float(value) for value in measurement[2]]
    if len(lengths) != 3 or min(lengths) <= 0:
        raise RuntimeError('Could not measure a valid workpiece bounding box.')
    max_length = max(lengths)
    if part.PartUnits == NXOpen.BasePart.Units.Inches:
        max_length *= 25.4
    box_points = [mcs_view_xyz(point) for point in measurement[0]]
    if len(box_points) < 8:
        raise RuntimeError('Could not obtain the workpiece bounding-box corners.')
    return (max(0.001, min(4.0, 72.0 / max_length)), box_points)


def mcs_view_create(part, sheet, model_view_name, origin, horizontal, normal, center, scale, angle=None):
    builder = part.DraftingViews.CreateBaseViewBuilder(None)
    try:
        builder.SelectModelView.SelectedView = part.ModelingViews.FindObject(model_view_name)
        builder.Style.ViewStyleHiddenLines.HiddenLine = False
        builder.Style.ViewStyleVisibleLines.VisibleColor = part.Colors.Find('White')
        if angle is not None:
            builder.Style.ViewStyleGeneral.AngleValue = float(angle)
        orientation = builder.Style.ViewStyleOrientation.Ovt
        orientation.AssociativeOrientation = False
        position = NXOpen.Point3d(*map(float, origin))
        update_option = NXOpen.SmartObject.UpdateOption.WithinModeling
        orientation.XDirection = part.Directions.CreateDirection(position, NXOpen.Vector3d(*map(float, horizontal)), update_option)
        orientation.NormalDirection = part.Directions.CreateDirection(position, NXOpen.Vector3d(*map(float, normal)), update_option)
        builder.Scale.Numerator = float(scale)
        builder.Scale.Denominator = 1.0
        sx, sy = (float(sheet.Length) / 210.0, float(sheet.Height) / 297.0)
        builder.Placement.Placement.SetValue(None, part.Views.WorkView, NXOpen.Point3d(float(center[0] * sx), float(center[1] * sy), 0.0))
        return builder.Commit()
    finally:
        builder.Destroy()


def mcs_view_create_note(part, view, text, position, name, text_size=None, text_font=None):
    builder = part.Annotations.CreateDraftingNoteBuilder(None)
    try:
        builder.Text.TextBlock.SetText([text])
        if text_size is not None:
            builder.Style.LetteringStyle.GeneralTextSize = float(text_size)
        if text_font is not None:
            builder.Style.LetteringStyle.GeneralTextFont = int(text_font)
        builder.Origin.Anchor = NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter
        builder.Origin.Origin.SetValue(None, view, position)
        note = builder.Commit()
    finally:
        builder.Destroy()
    origin = NXOpen.Annotations.Annotation.AssociativeOriginData()
    origin.OriginType = NXOpen.Annotations.AssociativeOriginType.RelativeToView
    origin.View = view
    note.SetAssociativeOrigin(origin, position)
    note.SetName(mcs_view_unique_name(name))
    if not note.HasAssociativeOrigin:
        raise RuntimeError('The {} label did not receive a RelativeToView association.'.format(text))
    return note


def mcs_view_add_xyz(part, uf, sheet, views, frame, scale, box_points, log):
    _, origin, axes = frame
    paper_unit = float(sheet.Length) / 210.0
    bold_font_index = part.Fonts.AddFont('blackfont', NXOpen.FontCollection.Type.Nx)
    if bold_font_index <= 0 or not part.Fonts.DoesFontExist(bold_font_index) or part.Fonts.GetFontType(bold_font_index) != NXOpen.FontCollection.Type.Nx:
        bold_font_index = 1
    label_distance = min(18.0, max(9.0, 13.0 * (scale / 0.47) ** 0.5)) * paper_unit
    text_size = min(4.5, max(2.4, 3.2 * label_distance / (13.0 * paper_unit))) * paper_unit
    center_mark_size = max(0.8 * paper_unit, 0.55 * text_size)
    text_size *= 1.35
    arrow_length = 3.0 * paper_unit
    edge_gap = 6.0 * paper_unit
    label_clearance = 10.0 * paper_unit
    reference_view_scale = float(scale)
    if not math.isfinite(reference_view_scale) or reference_view_scale <= 0.0:
        raise RuntimeError('The reference drafting-view scale must be positive.')
    hull_by_view = {view_name: projected_hull((mcs_view_project(uf, view, point) for point in box_points)) for view_name, view in views}
    made = []
    for view_name, view in views:
        view_scale = float(view.Scale)
        if not math.isfinite(view_scale) or view_scale <= 0.0:
            raise RuntimeError('The {} drafting view has an invalid scale: {}.'.format(view_name, view_scale))
        view_scale_ratio = view_scale / reference_view_scale
        view_text_size = text_size * view_scale_ratio
        view_label_clearance = label_clearance * view_scale_ratio
        base_x, base_y = mcs_view_project(uf, view, origin)
        center_position = NXOpen.Point3d(float(base_x), float(base_y), 0.0)
        center_note = mcs_view_create_note(part, view, '+', center_position, 'CAMV8_MCS_{}_ORIGIN_CENTER'.format(view_name), center_mark_size)
        made.append(center_note)
        view.ActivateForSketching()
        for label, axis in zip(('X', 'Y', 'Z'), axes):
            probe = uf.View.MapModelToDrawing(view.Tag, list(mcs_view_add(origin, axis, 10.0)))
            dx, dy = (float(probe[0]) - base_x, float(probe[1]) - base_y)
            projected_length = math.hypot(dx, dy)
            if projected_length <= max(1e-07, scale * 0.35):
                continue
            dx, dy = (dx / projected_length, dy / projected_length)
            edge_distance = ray_exit(hull_by_view[view_name], base_x, base_y, dx, dy)
            axis_length = max(0.0, edge_distance) + edge_gap if edge_distance is not None else max(label_distance, 12.0 * paper_unit)
            tip_x, tip_y = (base_x + axis_length * dx, base_y + axis_length * dy)
            arrow_a = (tip_x - arrow_length * dx - arrow_length * dy, tip_y - arrow_length * dy + arrow_length * dx)
            arrow_b = (tip_x - arrow_length * dx + arrow_length * dy, tip_y - arrow_length * dy - arrow_length * dx)
            if view_name == 'ISO':
                for start, end in (((base_x, base_y), (tip_x, tip_y)), ((tip_x, tip_y), arrow_a), ((tip_x, tip_y), arrow_b)):
                    part.Curves.CreateLine(NXOpen.Point3d(float(start[0]), float(start[1]), 0.0), NXOpen.Point3d(float(end[0]), float(end[1]), 0.0))
            note_position = NXOpen.Point3d(float(tip_x + (view_text_size + view_label_clearance) * dx), float(tip_y + (view_text_size + view_label_clearance) * dy), 0.0)
            note = mcs_view_create_note(part, view, label, note_position, 'CAMV8_MCS_{}_{}_RELATIVE_TO_VIEW'.format(view_name, label), view_text_size, bold_font_index)
            made.append(note)
    if not made:
        raise RuntimeError('No CAM-MCS XYZ labels could be projected into the test views.')
    sheet.ActivateForSketching()


def mcs_view_project(uf, view, point):
    mapped = uf.View.MapModelToDrawing(view.Tag, list(point))
    return (float(mapped[0]), float(mapped[1]))


def mcs_view_matching_views(sheet, uf, origin, axes):
    matches = {'XY': [], 'XZ': []}
    for view in sheet.GetDraftingViews():
        base = mcs_view_project(uf, view, origin)
        vectors = [(mcs_view_project(uf, view, mcs_view_add(origin, axis))[0] - base[0], mcs_view_project(uf, view, mcs_view_add(origin, axis))[1] - base[1]) for axis in axes]
        x_axis, y_axis, z_axis = vectors
        bounds = tuple((float(value) for value in uf.Draw.AskViewBorders(view.Tag)))

        def points_right(vector):
            return vector[0] > 1e-08 and abs(vector[1]) < abs(vector[0]) * 0.001

        def points_up(vector):
            return vector[1] > 1e-08 and abs(vector[0]) < abs(vector[1]) * 0.001
        if points_right(x_axis) and points_up(y_axis) and (math.hypot(*z_axis) < abs(x_axis[0]) * 0.001):
            matches['XY'].append((view, bounds))
        if points_right(x_axis) and points_up(z_axis) and (math.hypot(*y_axis) < abs(x_axis[0]) * 0.001):
            matches['XZ'].append((view, bounds))
    result = {}
    for plane, candidates in matches.items():
        if not candidates:
            raise RuntimeError('No CAM-MCS-aligned {} test view was found.'.format(plane))
        result[plane] = min(candidates, key=lambda item: item[1][0])
    return result


def mcs_view_drawing_point(uf, view, point):
    return mcs_view_project(uf, view, mcs_view_xyz(point))


def mcs_view_choose_geometry(part, sheet, uf, log):
    linear_edges = []
    circular_edges = []
    examined = 0
    for body in part.Bodies:
        if not body.IsSolidBody or body.IsBlanked:
            continue
        for edge in body.GetEdges():
            examined += 1
            if examined > MAX_MCS_DIMENSION_EDGES:
                raise RuntimeError('The lightweight annotation test is limited to 2000 solid edges.')
            if edge.SolidEdgeType == NXOpen.Edge.EdgeType.Circular:
                circular_edges.append(edge)
            elif edge.SolidEdgeType == NXOpen.Edge.EdgeType.Linear:
                p1, p2 = edge.GetVertices()
                linear_edges.append((edge, p1, p2))
    if not linear_edges and (not circular_edges):
        raise RuntimeError('No linear or circular solid edges were found for auto-dimensioning.')
    located = []
    for view in sheet.GetDraftingViews():
        bounds = tuple((float(value) for value in uf.Draw.AskViewBorders(view.Tag)))
        located.append((view, bounds))
    located.sort(key=lambda item: (round(item[1][0], 1), -item[1][3]))
    for view, bounds in located:
        best = None
        for edge, p1, p2 in linear_edges:
            a, b = (mcs_view_drawing_point(uf, view, p1), mcs_view_drawing_point(uf, view, p2))
            dx, dy = (abs(b[0] - a[0]), abs(b[1] - a[1]))
            span = max(dx, dy)
            if span < 1.0 or min(dx, dy) > 0.01 * span:
                continue
            if not all((bounds[0] - 0.5 <= p[0] <= bounds[2] + 0.5 and bounds[1] - 0.5 <= p[1] <= bounds[3] + 0.5 for p in (a, b))):
                continue
            if best is None or span > best[0]:
                best = (span, edge, p1, p2, a, b, dx >= dy)
        if best is not None:
            return (view, bounds, best, 'linear')
    if circular_edges and located:
        view, bounds = max(located, key=lambda item: min(item[1][2] - item[1][0], item[1][3] - item[1][1]))
        edge = max(circular_edges, key=lambda candidate: candidate.GetLength())
        p1, p2 = edge.GetVertices()
        a = mcs_view_drawing_point(uf, view, p1)
        return (view, bounds, (0.0, edge, p1, p2, a, a, True), 'diameter')
    raise RuntimeError('No suitable projected edge was found for auto-dimensioning.')


def mcs_view_endpoints(body, view, bounds, uf):
    edges = list(body.GetEdges())
    if len(edges) > MAX_MCS_DIMENSION_EDGES:
        raise RuntimeError('The reference body exceeds the 2000-edge annotation test limit.')
    points = []
    for edge in edges:
        if edge.SolidEdgeType != NXOpen.Edge.EdgeType.Linear:
            continue
        for model_point in edge.GetVertices():
            mapped = mcs_view_drawing_point(uf, view, model_point)
            if bounds[0] - 0.5 <= mapped[0] <= bounds[2] + 0.5 and bounds[1] - 0.5 <= mapped[1] <= bounds[3] + 0.5:
                points.append((edge, model_point, mapped))
    return points


def mcs_view_coordinate(model_point, origin, axis):
    values = mcs_view_xyz(model_point)
    return sum(((values[index] - origin[index]) * axis[index] for index in range(3)))


def mcs_view_enum_member(enum_type, names):
    for name in names:
        value = getattr(enum_type, name, None)
        if value is not None:
            return value
    available = [name for name in dir(enum_type) if not name.startswith('_')]
    raise RuntimeError('NX enum unavailable: {}.{}; available={}'.format(enum_type.__name__, '/'.join(names), available))


def mcs_view_single_side_style(dimension, resources, no_arrows):
    prefs = dimension.GetDimensionPreferences()
    resources.append(prefs)
    prefs.ArrowDisplay = no_arrows
    prefs.DisplayLineBetweenArrows = False
    prefs.ExtensionLineDisplay = NXOpen.Annotations.ExtensionLineDisplay.Second
    prefs.TextOrientation = NXOpen.Annotations.TextOrientation.Horizontal
    dimension.SetDimensionPreferences(prefs)
    dimension.NominalDecimalPlaces = 1
    dimension.MetricNominalDecimalPlaces = 1
    check = dimension.GetDimensionPreferences()
    resources.append(check)
    if check.ArrowDisplay != no_arrows or check.ExtensionLineDisplay != NXOpen.Annotations.ExtensionLineDisplay.Second or check.DisplayLineBetweenArrows:
        raise RuntimeError('NX did not accept the single-side dimension style.')


def mcs_view_create_zero_note(part, session, view, position, name, reference_dimension, resources, log):
    note_mark = session.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, name + ' zero marker')
    try:
        builder = part.Annotations.CreateDraftingNoteBuilder(None)
        try:
            builder.Text.TextBlock.SetText(['0.0'])
            builder.Origin.Anchor = NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter
            builder.Origin.Origin.SetValue(None, view, position)
            note = builder.Commit()
        finally:
            builder.Destroy()
        dimension_lettering = reference_dimension.GetLetteringPreferences()
        resources.append(dimension_lettering)
        note_lettering = note.GetLetteringPreferences()
        resources.append(note_lettering)
        note_lettering.SetGeneralText(dimension_lettering.GetDimensionText())
        note.SetLetteringPreferences(note_lettering)
        associative_origin = NXOpen.Annotations.Annotation.AssociativeOriginData()
        associative_origin.OriginType = NXOpen.Annotations.AssociativeOriginType.RelativeToView
        associative_origin.View = view
        note.SetAssociativeOrigin(associative_origin, position)
        note.SetName(mcs_view_unique_name(name))
        if not note.HasAssociativeOrigin:
            raise RuntimeError('Zero label did not receive a RelativeToView association.')
        if session.UpdateManager.DoUpdate(note_mark):
            raise RuntimeError('NX reported errors while creating a zero label.')
    except Exception:
        session.UndoToMark(note_mark, None)


def mcs_view_add_dimensions(part, sheet, uf, frame, log, session, zero):
    _, origin, axes = frame
    views = mcs_view_matching_views(sheet, uf, origin, axes)
    _, _, chosen, _ = mcs_view_choose_geometry(part, sheet, uf, log)
    body = chosen[1].GetBody()
    view_points = {}
    for plane in ('XY', 'XZ'):
        view, bounds = views[plane]
        candidates = mcs_view_endpoints(body, view, bounds, uf)
        if not candidates:
            raise RuntimeError('No visible linear edge endpoints in the {} view.'.format(plane))
        view_points[plane] = (view, bounds, candidates)
    plans = []
    for axis_index, plane, horizontal in ((0, 'XY', True), (1, 'XY', False), (2, 'XZ', False)):
        view, bounds, candidates = view_points[plane]
        low = min(candidates, key=lambda item: mcs_view_coordinate(item[1], origin, axes[axis_index]))
        high = max(candidates, key=lambda item: mcs_view_coordinate(item[1], origin, axes[axis_index]))
        low_value = mcs_view_coordinate(low[1], origin, axes[axis_index])
        high_value = mcs_view_coordinate(high[1], origin, axes[axis_index])
        if high_value - low_value <= 1e-05:
            raise RuntimeError('{} direction has no measurable span.'.format('XYZ'[axis_index]))
        plans.append((axis_index, horizontal, view, bounds, low, low_value, high, high_value))
    no_arrows = mcs_view_enum_member(NXOpen.Annotations.ArrowDisplay, ('NotSet', 'None', 'NoneOption', 'None_'))
    resources = []
    mark = session.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, 'V12-style XYZ auto-dimension test')
    made = []
    zero_notes = []
    try:
        for axis_index, horizontal, view, bounds, low, low_value, high, high_value in plans:
            axis_name = 'XYZ'[axis_index]
            origin_2d = mcs_view_project(uf, view, origin)
            view_width = bounds[2] - bounds[0]
            view_height = bounds[3] - bounds[1]
            dynamic_gap = max(4.0, min(view_width, view_height) * 0.12)
            if part.PartUnits != NXOpen.BasePart.Units.Millimeters:
                dynamic_gap /= 25.4

            if horizontal:
                dimension_sides = (('BOTTOM', bounds[1] - dynamic_gap), ('TOP', bounds[3] + dynamic_gap))
                zero_positions = [(side_name, NXOpen.Point3d(float(origin_2d[0]), float(side_position), 0.0)) for side_name, side_position in dimension_sides]
            else:
                dimension_sides = (('LEFT', bounds[0] - dynamic_gap), ('RIGHT', bounds[2] + dynamic_gap))
                zero_positions = [(side_name, NXOpen.Point3d(float(side_position), float(origin_2d[1]), 0.0)) for side_name, side_position in dimension_sides]
            if axis_name != 'Z':
                for side_name, position in zero_positions:
                    zero_notes.append((view, position, 'VIEWXYZ_TEST_{}_ZERO_{}'.format(axis_name, side_name)))
            endpoints = [(low[0], low[1], low[2], low_value), (high[0], high[1], high[2], high_value)]
            if axis_name == 'Z' and (not any((abs(item[3]) <= 1e-05 for item in endpoints))):
                zero_endpoint = next((item for item in view_points['XZ'][2] if abs(mcs_view_coordinate(item[1], origin, axes[2])) <= 1e-05), None)
                if zero_endpoint is None:
                    raise RuntimeError('No real solid-edge endpoint lies on MCS Z=0 for a native zero dimension.')
                endpoints.append((zero_endpoint[0], zero_endpoint[1], zero_endpoint[2], 0.0))
            for edge, model_point, mapped, value in endpoints:
                if abs(value) <= 1e-05 and axis_name != 'Z':
                    continue
                endpoint_name = 'LOW' if abs(value - low_value) <= 1e-05 else 'HIGH' if abs(value - high_value) <= 1e-05 else 'ZERO'
                for side_name, side_position in dimension_sides:
                    data = part.Annotations.NewDimensionData()
                    resources.append(data)
                    for index, obj, pick in ((1, zero, NXOpen.Point3d(*map(float, origin))), (2, edge, model_point)):
                        association = part.Annotations.NewAssociativity()
                        resources.append(association)
                        association.FirstObject = obj
                        association.ObjectView = view
                        association.PointOption = NXOpen.Annotations.AssociativityPointOption.Control
                        association.PickPoint = pick
                        data.SetAssociativity(index, [association])
                    dim_position = NXOpen.Point3d(float(mapped[0] if horizontal else side_position), float(side_position if horizontal else mapped[1]), 0.0)
                    factory = part.Dimensions.CreateHorizontalDimension if horizontal else part.Dimensions.CreateVerticalDimension
                    dimension = factory(data, dim_position)
                    dimension.SetName(mcs_view_unique_name('VIEWXYZ_TEST_{}_{}_{}'.format(axis_name, endpoint_name, side_name)))
                    mcs_view_single_side_style(dimension, resources, no_arrows)
                    if value < 0:
                        appended = dimension.GetAppendedText()
                        resources.append(appended)
                        appended.SetBeforeText(['-'])
                        dimension.SetAppendedText(appended)
                    else:
                        prefs = dimension.GetDimensionPreferences()
                        resources.append(prefs)
                        prefs.ExtensionLineDisplay = NXOpen.Annotations.ExtensionLineDisplay.Two
                        dimension.SetDimensionPreferences(prefs)
                    original_text_origin = dimension.AnnotationOrigin
                    text_offset = 0.2 * dynamic_gap
                    if axis_name == 'X':
                        vertical_offset = text_offset if side_name == 'TOP' else -text_offset
                        dimension.AnnotationOrigin = NXOpen.Point3d(float(original_text_origin.X), float(original_text_origin.Y + vertical_offset), float(original_text_origin.Z))
                    elif axis_name in ('Y', 'Z'):
                        horizontal_offset = text_offset if side_name == 'RIGHT' else -text_offset
                        dimension.AnnotationOrigin = NXOpen.Point3d(float(original_text_origin.X + horizontal_offset), float(original_text_origin.Y), float(original_text_origin.Z))
                    made.append((dimension, value, view, axis_name))
        if session.UpdateManager.DoUpdate(mark):
            raise RuntimeError('NX reported errors while updating V12-style dimensions.')
        for view, position, note_name in zero_notes:
            reference = next((dimension for dimension, _, candidate, _ in made if candidate.Tag == view.Tag), None)
            if reference is not None:
                mcs_view_create_zero_note(part, session, view, position, note_name, reference, resources, log)
    except Exception:
        session.UndoToMark(mark, None)
        raise
    finally:
        for resource in reversed(resources):
            try:
                resource.Dispose()
            except (AttributeError, NXOpen.NXException):
                pass


class MCSViewLog:
    def __init__(self, listing):
        self.listing = listing

    def write(self, message):
        self.listing.WriteLine(str(message))


def add_model_views(part, uf, sheet, listing, settings):
    bodies = [body for body in part.Bodies if body.IsSolidBody and not body.IsBlanked]
    if not bodies:
        listing.WriteLine("No visible solid body; skipping CAM views and dimensions.")
        return ""

    global MCS_VIEW_RUN_ID
    MCS_VIEW_RUN_ID = "{}_{}".format(sheet.Name[-20:], datetime.now().strftime("%H%M%S_%f"))
    log = MCSViewLog(listing)
    session = NXOpen.Session.GetSession()
    if len(MCS_FRAMES) == 1:
        frame = next(iter(MCS_FRAMES.values()))
    else:
        frame = mcs_view_cam_mcs(part, log)

    scale, box_points = mcs_view_choose_scale(part, session)
    _, origin, (xm, ym, zm) = frame
    xy_center = (float(settings["view_lower_x"]), float(settings["view_lower_y"]))
    xz_center = (float(settings["view_upper_x"]), float(settings["view_upper_y"]))
    iso_center = (float(settings["view_iso_x"]), float(settings["view_iso_y"]))

    xy_view = mcs_view_create(part, sheet, "Top", origin, xm, zm, xy_center, scale)
    xz_view = mcs_view_create(part, sheet, "Front", origin, xm, tuple(-value for value in ym), xz_center, scale)
    tilt = math.radians(35.0)
    azimuth = math.radians(45.0)
    normal = tuple(math.cos(tilt) * z + math.sin(tilt) * (math.cos(azimuth) * x - math.sin(azimuth) * y) for x, y, z in zip(xm, ym, zm))
    projection = mcs_view_dot(xm, normal)
    horizontal = mcs_view_unit(tuple(x - projection * n for x, n in zip(xm, normal)))
    iso_view = mcs_view_create(part, sheet, "Isometric", origin, horizontal, normal, iso_center, scale, angle=-35.0)
    views = (("XY", xy_view), ("XZ", xz_view), ("ISO", iso_view))

    part.DrawingSheets.RefreshCurrentSheet()
    zero = part.Points.CreatePoint(NXOpen.Point3d(*map(float, origin)))
    zero.SetName(mcs_view_unique_name("CAM_MCS_ORIGIN_ANCHOR"))

    mcs_view_add_dimensions(part, sheet, uf, frame, log, session, zero)
    mcs_view_add_xyz(part, uf, sheet, views, frame, scale, box_points, log)
    return "AUTO"


CONFIG_PATH = Path.home() / "Documents" / "Codex" / "2026-11-10" / "kai" / "outputs" / "CAM_Sheet_Config.json"

DEFAULTS = {
    "programmer": "", "xy": "", "z": "",
    "xy_auto": "true", "z_auto": "true",
    "origin_settings_version": "2",
    "mode": "新建程序单", "output_dir": "",
    "template_dir": str(CONFIG_PATH.parent),
    "font_family": "宋体",
    "heading_size": "1.8",
    "table_spacing": "1.0",
    "table_aspect": "0.5",
    "table_row_height": "6.0",
    "upper_size": "2.5",
    "company_size": "5.0",
    "company_name": "广东英维克技术有限公司",
    "length_allowance": "2",
    "stock_mode": "auto",
    "stock_size": "66*56*14",
    "stock_auto_size": "自动识别失败",
    "quantity": "1",
    "material": "铝件",
    "clamp_methods": "虎钳一序",
    "clamp_depth": "5",
    "clamp_depth_suppressed": "false",
    "layout": "{}",
    "view_upper_x": "48", "view_upper_y": "238",
    "view_lower_x": "48", "view_lower_y": "193",
    "view_iso_x": "157", "view_iso_y": "200",
}

DIALOG = r"""
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$cfg = $env:CODEX_CAM_SHEET_SETTINGS | ConvertFrom-Json
$f = New-Object Windows.Forms.Form
$f.Text = 'NX 程序单设置'
$f.ClientSize = New-Object Drawing.Size(900, 750)
$f.AutoScroll = $true
$f.BackColor = [Drawing.Color]::FromArgb(245,248,252)
$f.Font = New-Object Drawing.Font('Microsoft YaHei UI', 9)
$f.FormBorderStyle = 'FixedDialog'
$f.MaximizeBox = $false
$banner = New-Object Windows.Forms.Label
$banner.Text = '  CAM 程序单  |  出单设置'
$banner.Font = New-Object Drawing.Font('Microsoft YaHei UI', 15, [Drawing.FontStyle]::Bold)
$banner.ForeColor = [Drawing.Color]::White
$banner.BackColor = [Drawing.Color]::FromArgb(25,48,78)
$banner.TextAlign = 'MiddleLeft'
$banner.SetBounds(0,0,900,54)
$f.Controls.Add($banner)
$divider = New-Object Windows.Forms.Panel
$divider.BackColor = [Drawing.Color]::FromArgb(190,200,212)
$divider.SetBounds(450,60,2,675)
$f.Controls.Add($divider)
$f.StartPosition = 'CenterScreen'
$f.TopMost = $true

$labels = [ordered]@{programmer='编程人员';company_name='公司名称';xy='XY 取数说明';z='Z 原点说明';mode='出单方式';output_dir='输出目录（预留）';template_dir='模板目录（预留）';font_family='程序单字体';company_size='公司名称字号 (mm)';upper_size='上方信息字号 (mm)';heading_size='下方参数表字号 (mm)';length_allowance='刀长加量 (mm)';table_aspect='参数表文本宽高比';table_spacing='参数字间距系数';table_row_height='参数表行高 (mm)'}
$controls = @{}
$originChecks = @{}
$i = 0
foreach ($key in $labels.Keys) {
    if ($i -in @(0,5)) {
        $section = New-Object Windows.Forms.Label
        $section.Text = @{0='01  程序单信息';5='02  文件目录'}[$i]
        $section.ForeColor = [Drawing.Color]::FromArgb(23,91,160)
        $section.SetBounds(20,(65 + 32*$i + 26*[int]($i -ge 5) + 26*[int]($i -ge 7)),420,23)
        $f.Controls.Add($section)
    }
    $y = 90 + 32*$i + 26*[int]($i -ge 5) + 26*[int]($i -ge 7)
    $i++
    $label = New-Object Windows.Forms.Label
    $label.Text = $labels[$key]
    $label.SetBounds(12,$y,125,24)
    $f.Controls.Add($label)
    if ($key -eq 'mode') {
        $box = New-Object Windows.Forms.ComboBox
        $box.DropDownStyle = 'DropDownList'
        [void]$box.Items.AddRange([object[]]@('新建程序单','刷新本次会话的已有程序单'))
        $box.SelectedIndex = 0
    } elseif ($key -eq 'font_family') {
        $box = New-Object Windows.Forms.ComboBox
        $box.DropDownStyle = 'DropDown'
        $fonts = New-Object Drawing.Text.InstalledFontCollection
        try {
            $names = @('宋体', '仿宋') + @($fonts.Families | ForEach-Object Name)
            $box.Items.AddRange([object[]]($names | Sort-Object -Unique))
        } finally { $fonts.Dispose() }
    } else {
        $box = New-Object Windows.Forms.TextBox
    }
    $box.Text = [string]$cfg.$key
    $box.SetBounds(140,$y,270,27)
    if ($key -in @('xy','z')) {
        $box.Width = 180
        $auto = New-Object Windows.Forms.CheckBox
        $auto.Text = '自动识别'
        $auto.SetBounds(330,$y,80,27)
        $auto.Checked = [string]$cfg.($key+'_auto') -eq 'true'
        $box.ReadOnly = $auto.Checked
        $auto.Tag = @{box=$box; value=[string]$cfg.($key+'_detected')}
        $auto.Add_CheckedChanged({
            param($sender,$eventArgs)
            $sender.Tag.box.ReadOnly = $sender.Checked
            if ($sender.Checked) { $sender.Tag.box.Text = $sender.Tag.value }
        })
        $f.Controls.Add($auto)
        $originChecks[$key] = $auto
    }
    if ($key -in @('output_dir','template_dir')) {
        $box.Width = 180
        $browse = New-Object Windows.Forms.Button
        $browse.Text = '浏览…'
        $browse.Tag = $box
        $browse.SetBounds(325,$y,85,27)
        $browse.Add_Click({
            param($sender, $eventArgs)
            $target = $sender.Tag
            $picker = New-Object Windows.Forms.FolderBrowserDialog
            try {
                $picker.Description = '选择目录'
                if ([IO.Directory]::Exists($target.Text)) { $picker.SelectedPath = $target.Text }
                if ($picker.ShowDialog($sender.FindForm()) -eq 'OK') { $target.Text = $picker.SelectedPath }
            } finally { $picker.Dispose() }
        })
        $f.Controls.Add($browse)
    }
    $f.Controls.Add($box)
    $controls[$key] = $box
}

foreach ($entry in @(@('确定','OK',700), @('取消','Cancel',790))) {
    $button = New-Object Windows.Forms.Button
    $button.Text = $entry[0]
    $button.DialogResult = $entry[1]
    $button.SetBounds($entry[2],707,80,28)
    if ($entry[1] -eq 'OK') {
        $button.FlatStyle = 'Flat'
        $button.BackColor = [Drawing.Color]::FromArgb(23,91,160)
        $button.ForeColor = [Drawing.Color]::White
    }
    $f.Controls.Add($button)
    if ($entry[1] -eq 'OK') { $f.AcceptButton = $button } else { $f.CancelButton = $button }
}

$layoutJson = [string]$cfg.layout
$layoutButton = New-Object Windows.Forms.Button
$layoutButton.Text = '文字位置…'
$layoutButton.SetBounds(20,707,125,28)
$layoutButton.Add_Click({
    $dialog = New-Object Windows.Forms.Form
    $dialog.Text = '文字位置（毫米）：X 向右，Y 向上'
    $dialog.ClientSize = New-Object Drawing.Size(510,540)
    $dialog.StartPosition = 'CenterParent'
    $grid = New-Object Windows.Forms.DataGridView
    $grid.SetBounds(12,12,486,478)
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AutoSizeColumnsMode = 'Fill'
    [void]$grid.Columns.Add('item','项目')
    [void]$grid.Columns.Add('dx','X 偏移')
    [void]$grid.Columns.Add('dy','Y 偏移')
    $grid.Columns[0].ReadOnly = $true
    $saved = $script:layoutJson | ConvertFrom-Json
    foreach ($name in ($env:CODEX_CAM_LAYOUT_KEYS | ConvertFrom-Json)) {
        $entry = $saved.PSObject.Properties[$name]
        $pair = if ($null -ne $entry) { $entry.Value } else { @(0,0) }
        [void]$grid.Rows.Add([object[]]@($name,$pair[0],$pair[1]))
    }
    $dialog.Controls.Add($grid)
    $ok = New-Object Windows.Forms.Button
    $ok.Text = '应用'
    $ok.SetBounds(316,501,85,28)
    $ok.Add_Click({
        $grid.EndEdit()
        $offsets = @{}
        foreach ($row in $grid.Rows) {
            $x = 0.0; $y = 0.0
            if (-not [double]::TryParse([string]$row.Cells[1].Value,[ref]$x) -or
                -not [double]::TryParse([string]$row.Cells[2].Value,[ref]$y)) {
                return
            }
            if ($x -ne 0 -or $y -ne 0) { $offsets[[string]$row.Cells[0].Value] = @($x,$y) }
        }
        $dialog.Tag = $offsets | ConvertTo-Json -Compress
        $dialog.Close()
    }.GetNewClosure())
    $dialog.Controls.Add($ok)
    $cancel = New-Object Windows.Forms.Button
    $cancel.Text = '取消'
    $cancel.DialogResult = 'Cancel'
    $cancel.SetBounds(412,501,85,28)
    $dialog.CancelButton = $cancel
    $dialog.Controls.Add($cancel)
    [void]$dialog.ShowDialog($f)
    if ($dialog.Tag) { $script:layoutJson = [string]$dialog.Tag }
    $dialog.Dispose()
})
$f.Controls.Add($layoutButton)

$script:viewPositions = @{}
foreach ($key in @('view_upper_x','view_upper_y','view_lower_x','view_lower_y','view_iso_x','view_iso_y')) {
    $script:viewPositions[$key] = [string]$cfg.$key
}
$positionButton = New-Object Windows.Forms.Button
$positionButton.Text = '三视图位置…'
$positionButton.SetBounds(155,707,140,28)
$positionButton.Add_Click({
    $d = New-Object Windows.Forms.Form
    $d.Text = '三视图位置'
    $d.ClientSize = New-Object Drawing.Size(490,285)
    $d.StartPosition = 'CenterParent'
    $d.Font = $f.Font
    $d.FormBorderStyle = 'FixedDialog'
    $d.MaximizeBox = $false
    $hint = New-Object Windows.Forms.Label
    $hint.Text = '图纸左下角为原点，单位 mm；X 向右、Y 向上。'
    $hint.SetBounds(15,12,460,35)
    $d.Controls.Add($hint)
    $fields = @{}
    $rows = @(@('左上视图','view_upper',48,238),@('左下视图','view_lower',48,193),@('右侧等轴视图','view_iso',157,200))
    $r = 0
    foreach ($row in $rows) {
        $label = New-Object Windows.Forms.Label
        $label.Text = $row[0]
        $label.SetBounds(15,(72+45*$r),115,25)
        $d.Controls.Add($label)
        $j = 0
        foreach ($axis in @('x','y')) {
            $caption = New-Object Windows.Forms.Label
            $caption.Text = $axis.ToUpper()
            $caption.SetBounds((135+165*$j),(72+45*$r),25,25)
            $d.Controls.Add($caption)
            $key = $row[1]+'_'+$axis
            $n = New-Object Windows.Forms.NumericUpDown
            $n.DecimalPlaces = 2
            $n.Minimum = 0
            $n.Maximum = if ($axis -eq 'x') { 210 } else { 297 }
            $n.Value = [decimal]::Parse($script:viewPositions[$key],[Globalization.CultureInfo]::InvariantCulture)
            $n.SetBounds((165+165*$j),(70+45*$r),120,27)
            $d.Controls.Add($n)
            $fields[$key] = $n
            $j++
        }
        $r++
    }
    $ok = New-Object Windows.Forms.Button
    $ok.Text = '确定'
    $ok.DialogResult = 'OK'
    $ok.SetBounds(290,232,80,30)
    $d.Controls.Add($ok)
    $d.AcceptButton = $ok
    $cancel = New-Object Windows.Forms.Button
    $cancel.Text = '取消'
    $cancel.DialogResult = 'Cancel'
    $cancel.SetBounds(385,232,80,30)
    $d.Controls.Add($cancel)
    $d.CancelButton = $cancel
    if ($d.ShowDialog($f) -eq 'OK') {
        foreach ($key in $fields.Keys) {
            $script:viewPositions[$key] = $fields[$key].Value.ToString([Globalization.CultureInfo]::InvariantCulture)
        }
    }
    $d.Dispose()
})
$f.Controls.Add($positionButton)

$stockSection = New-Object Windows.Forms.Label
$stockSection.Text = '04  毛坯与加工要求'
$stockSection.ForeColor = [Drawing.Color]::FromArgb(23,91,160)
$stockSection.SetBounds(470,65,410,23)
$f.Controls.Add($stockSection)

$stockModeLabel = New-Object Windows.Forms.Label
$stockModeLabel.Text = '毛坯尺寸模式'
$stockModeLabel.SetBounds(470,100,110,24)
$f.Controls.Add($stockModeLabel)

$stockMode = New-Object Windows.Forms.ComboBox
$stockMode.DropDownStyle = 'DropDownList'
[void]$stockMode.Items.AddRange([object[]]@('自动读取零件尺寸','手动输入毛坯尺寸'))
$stockMode.SelectedIndex = if ([string]$cfg.stock_mode -eq 'manual') { 1 } else { 0 }

$stockSizeLabel = New-Object Windows.Forms.Label
$stockSizeLabel.Text = '毛坯尺寸 XYZ (mm)'
$stockSizeLabel.SetBounds(470,140,110,24)
$f.Controls.Add($stockSizeLabel)

$stockSize = New-Object Windows.Forms.TextBox
$manualStockSize = [string]$cfg.stock_size
$autoStockSize = [string]$cfg.stock_auto_size

if (-not $manualStockSize -or $manualStockSize -eq '自动识别失败') {
    $manualStockSize = $autoStockSize
}
$stockSize.Text = if ($stockMode.SelectedIndex -eq 0) { $autoStockSize } else { $manualStockSize }
$stockSize.ReadOnly = $stockMode.SelectedIndex -eq 0
$stockSize.SetBounds(590,136,290,27)

$stockMode.Tag = @{sizeBox=$stockSize; autoSize=$autoStockSize; manualSize=$manualStockSize}
$stockSize.Tag = $stockMode

$stockSize.Add_TextChanged({
    param($sender,$eventArgs)
    $owner = $sender.Tag
    if ($owner.SelectedIndex -eq 1) { $owner.Tag.manualSize = $sender.Text }
}.GetNewClosure())

$stockMode.Add_SelectedIndexChanged({
    param($sender,$eventArgs)
    $box = $sender.Tag.sizeBox
    if ($sender.SelectedIndex -eq 0) {
        $box.ReadOnly = $true
        $box.Text = $sender.Tag.autoSize
    } else {
        $box.ReadOnly = $false
        if (-not $sender.Tag.manualSize -or $sender.Tag.manualSize -eq '自动识别失败') {
            $sender.Tag.manualSize = $sender.Tag.autoSize
        }
        $box.Text = $sender.Tag.manualSize
    }
}.GetNewClosure())

$stockMode.SetBounds(590,96,290,27)
$f.Controls.Add($stockMode)
$f.Controls.Add($stockSize)

$stockHint = New-Object Windows.Forms.Label
$stockHint.Text = '自动：按 MCS 尺寸单侧加 3 mm；手动模式支持在此基础上直接修改。'
$stockHint.SetBounds(590,166,290,34)
$f.Controls.Add($stockHint)

$quantityLabel = New-Object Windows.Forms.Label
$quantityLabel.Text = '需求数量'
$quantityLabel.SetBounds(470,215,110,24)
$f.Controls.Add($quantityLabel)
$quantity = New-Object Windows.Forms.NumericUpDown
$quantity.DecimalPlaces = 0
$quantity.Minimum = 1
$quantity.Maximum = 100000
$quantityValue = 1
[void][int]::TryParse([string]$cfg.quantity,[ref]$quantityValue)
$quantity.Value = [Math]::Min(100000,[Math]::Max(1,$quantityValue))
$quantity.SetBounds(590,211,150,27)
$f.Controls.Add($quantity)

$materialLabel = New-Object Windows.Forms.Label
$materialLabel.Text = '材质'
$materialLabel.SetBounds(470,255,110,24)
$f.Controls.Add($materialLabel)
$material = New-Object Windows.Forms.ComboBox
$material.DropDownStyle = 'DropDown'
[void]$material.Items.AddRange([object[]]@('铝件','钢件','不锈钢','铜件','塑料'))
$material.Text = [string]$cfg.material
$material.SetBounds(590,251,290,27)
$f.Controls.Add($material)

$clampLabel = New-Object Windows.Forms.Label
$clampLabel.Text = '夹持方式'
$clampLabel.SetBounds(470,295,110,24)
$f.Controls.Add($clampLabel)
$savedMethods = [string]$cfg.clamp_methods -split '[\r\n;；]+'
$clampChoices = @('虎钳一序','虎钳二序','吸盘一序','吸盘二序','胶水固定','AB胶固定','压块固定','其他')
$clampMethod = New-Object Windows.Forms.ComboBox
$clampMethod.DropDownStyle = 'DropDown'
[void]$clampMethod.Items.AddRange([object[]]$clampChoices)
if ($savedMethods.Count -gt 0) { $clampMethod.Text = $savedMethods[0].Trim() }
$clampMethod.SetBounds(590,291,290,27)

$clampMethod.Add_SelectedIndexChanged({
    param($sender,$eventArgs)
    $val = $sender.Text
    if ($val -match '吸盘|胶水|压块') {
        [void][Windows.Forms.MessageBox]::Show(
            "当前选择的是 [$val]，建议检查并关闭不必要的夹持深度（可勾选右侧'不输出深度'）。",
            "夹持方式提示",
            [Windows.Forms.MessageBoxButtons]::OK,
            [Windows.Forms.MessageBoxIcon]::Information
        )
    }
})
$f.Controls.Add($clampMethod)

$clampDepthLabel = New-Object Windows.Forms.Label
$clampDepthLabel.Text = '夹持深度 (mm)'
$clampDepthLabel.SetBounds(470,335,110,24)
$f.Controls.Add($clampDepthLabel)
$clampDepth = New-Object Windows.Forms.TextBox
$clampDepth.Text = [string]$cfg.clamp_depth
$clampDepth.SetBounds(590,331,150,27)
$f.Controls.Add($clampDepth)

$clampDepthSuppressed = New-Object Windows.Forms.CheckBox
$clampDepthSuppressed.Text = '不输出深度'
$clampDepthSuppressed.Checked = [string]$cfg.clamp_depth_suppressed -eq 'true'
$clampDepthSuppressed.Tag = $clampDepth
$clampDepthSuppressed.SetBounds(750,331,130,27)
$clampDepthSuppressed.Add_CheckedChanged({
    param($sender,$eventArgs)
    $sender.Tag.Enabled = -not $sender.Checked
})
$clampDepth.Enabled = -not $clampDepthSuppressed.Checked
$f.Controls.Add($clampDepthSuppressed)

$setupHint = New-Object Windows.Forms.Label
$setupHint.Text = '常用方式可下拉选择，若选择吸盘/压块等会给出深度关闭提示。'
$setupHint.SetBounds(470,375,410,40)
$f.Controls.Add($setupHint)

if ($f.ShowDialog() -eq 'OK') {
    $result = @{}
    foreach ($key in $labels.Keys) { $result[$key] = $controls[$key].Text.Trim() }
    foreach ($key in $originChecks.Keys) {
        $result[$key+'_auto'] = $originChecks[$key].Checked.ToString().ToLowerInvariant()
    }
    $result["layout"] = $script:layoutJson
    foreach ($key in $script:viewPositions.Keys) { $result[$key] = $script:viewPositions[$key] }
    $result["stock_mode"] = if ($stockMode.SelectedIndex -eq 0) { "auto" } else { "manual" }
    $result["stock_size"] = [string]$stockMode.Tag.manualSize
    if ($stockMode.SelectedIndex -eq 1) { $result["stock_size"] = $stockSize.Text.Trim() }
    $result["quantity"] = [string]$quantity.Value
    $result["material"] = $material.Text.Trim()
    $result["clamp_methods"] = $clampMethod.Text.Trim()
    $result["clamp_depth"] = $clampDepth.Text.Trim()
    $result["clamp_depth_suppressed"] = $clampDepthSuppressed.Checked.ToString().ToLowerInvariant()
    [Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
    $result | ConvertTo-Json -Compress
}
$f.Dispose()
"""


class RunLogger:
    def __init__(self):
        directory = os.path.join(tempfile.gettempdir(), "NX_CAM_ProgramSheet")
        os.makedirs(directory, exist_ok=True)
        filename = "CAM_V31_{}.log".format(datetime.now().strftime("%Y%m%d_%H%M%S_%f"))
        self.path = os.path.join(directory, filename)
        self._file = open(self.path, "w", encoding="utf-8")
        self.WriteLine("Run log: " + self.path)

    def WriteLine(self, message):
        self._file.write(str(message) + "\n")
        self._file.flush()

    def close(self):
        if not self._file.closed:
            self._file.close()


def get_settings(part, listing):
    settings = DEFAULTS.copy()
    if CONFIG_PATH.is_file():
        saved = json.loads(CONFIG_PATH.read_text(encoding="utf-8-sig"))
        settings.update({key: str(saved[key]) for key in settings if key in saved})

    group = selected_program_group(part, listing)
    if not group:
        return None
    operation_items = list(grouped_operations(group))
    for _, operation in operation_items:
        operation_geometry(operation, NXOpen.CAM.CAMSetup.View, listing)

    probe = dict(settings, xy_auto="true", z_auto="true")
    detected = resolve_origin_description(part, probe, listing)
    for key in ("xy", "z"):
        settings[key + "_detected"] = detected[key]
        if settings[key + "_auto"] == "true":
            settings[key] = detected[key]
        elif settings[key] == "自动判断":
            settings[key] = detected[key]

    part_dimensions = mcs_workpiece_dimensions_mm(part, listing)
    if part_dimensions is None:
        auto_stock_size = "自动识别失败，请切换手动输入"
    else:
        auto_stock_size = format_stock_size(tuple(value + 6.0 for value in part_dimensions))
    settings["stock_auto_size"] = auto_stock_size

    env = os.environ.copy()
    env["CODEX_CAM_SHEET_SETTINGS"] = json.dumps(settings)
    env["CODEX_CAM_LAYOUT_KEYS"] = json.dumps(LAYOUT_KEYS, ensure_ascii=False)

    descriptor, script_path = tempfile.mkstemp(prefix="cam_sheet_settings_", suffix=".ps1")
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as script_file:
            script_file.write(DIALOG)
        env["CODEX_CAM_DIALOG_PATH"] = script_path
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-STA", "-Command",
             "Invoke-Expression ([IO.File]::ReadAllText($env:CODEX_CAM_DIALOG_PATH, [Text.Encoding]::UTF8))"],
            env=env, capture_output=True, encoding="utf-8", check=True,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
    finally:
        os.remove(script_path)
    if not result.stdout.strip():
        return None
    settings = json.loads(result.stdout)
    settings["stock_auto_size"] = auto_stock_size
    settings["origin_settings_version"] = "2"
    for key in ("xy", "z"):
        if not settings[key].strip() or settings[key].strip() == "自动判断":
            settings[key + "_auto"] = "true"
        if settings[key + "_auto"] == "true":
            settings[key] = detected[key]
        else:
            settings[key] = settings[key].strip()

    settings["font_family"] = settings["font_family"].strip() or "宋体"
    heading_size = float(settings["heading_size"] or "1.8")
    settings["heading_size"] = str(heading_size)
    upper_size = float(settings["upper_size"] or "2.5")
    settings["upper_size"] = str(upper_size)
    company_size = float(settings["company_size"] or "5.0")
    settings["company_size"] = str(company_size)
    settings["length_allowance"] = settings["length_allowance"].strip() or "2"
    settings["stock_mode"] = settings.get("stock_mode", "auto").strip().lower()

    if settings["stock_mode"] == "auto":
        if auto_stock_size.startswith("自动识别失败"):
            raise ValueError(auto_stock_size + "。")
        stock_dimensions = parse_stock_size(auto_stock_size)
    else:
        stock_dimensions = parse_stock_size(settings.get("stock_size", ""))
        settings["stock_size"] = format_stock_size(stock_dimensions)
    settings["stock_size_resolved"] = format_stock_size(stock_dimensions)

    try:
        quantity = int(settings.get("quantity", "1"))
    except (TypeError, ValueError):
        quantity = 1
    settings["quantity"] = str(quantity)
    settings["material"] = settings.get("material", "").strip() or "铝件"

    methods = [value.strip() for value in re.split(
        r"[\r\n;；]+", settings.get("clamp_methods", "")) if value.strip()]
    settings["clamp_methods"] = normalize_clamp_method(methods[0] if methods else "虎钳一序")
    settings["clamp_depth_suppressed"] = str(settings.get("clamp_depth_suppressed", "false")).strip().lower()

    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")
    return settings


def add_sheet_border(part, sheet):
    sheet.ActivateForSketching()
    x0, y0 = 2.0 * float(sheet.Length) / 210.0, 1.0
    x1, y1 = 208.0 * float(sheet.Length) / 210.0, float(sheet.Height) - 1.0
    corners = [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]
    for start, end in zip(corners, corners[1:] + corners[:1]):
        part.Curves.CreateLine(NXOpen.Point3d(*start, 0.0), NXOpen.Point3d(*end, 0.0))


def main():
    global FONT_INDEX, TABLE_CHAR_SPACING, TABLE_ASPECT, RUN_LOGGER
    FONT_WARNINGS.clear()
    MCS_FRAMES.clear()
    WORKPIECE_BOX_CACHE.clear()
    GROUPED_OPERATIONS_CACHE.clear()
    OPERATION_GEOMETRY_CACHE.clear()
    session = NXOpen.Session.GetSession()
    listing = RunLogger()
    RUN_LOGGER = listing
    mark = None
    successful = False
    try:
        part = session.Parts.Work
        displayed = session.Parts.Display
        if part is None or displayed is None or part.Tag != displayed.Tag:
            raise RuntimeError("Open the CAM part as the work and display part.")

        settings = get_settings(part, listing)
        if settings is None:
            return
        LAYOUT.clear()
        LAYOUT.update({key: value for key, value in json.loads(settings["layout"]).items()
                       if key in LAYOUT_KEYS})
        TABLE_CHAR_SPACING = float(settings["table_spacing"])
        TABLE_ASPECT = float(settings["table_aspect"])
        settings = resolve_origin_description(part, settings, listing)
        rows_per_page = rows_per_page_for_settings(settings)
        result = collect(part, listing, float(settings["length_allowance"]))
        if not result:
            return
        group_name, (rows, total) = result
        session.ApplicationSwitchImmediate("UG_APP_DRAFTING")
        mark = session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "CAM tool sheets")
        uf = NXOpen.UF.UFSession.GetUFSession()
        drawing = NXOpen.Drawings.DrawingSheet
        stamp = datetime.now().strftime("%H%M%S_%f")
        cache = sys.__dict__.setdefault("_codex_cam_shortremarks_v8_refresh", {})
        key = (part.Tag, part.FullPath, group_name)
        signature = (settings["table_row_height"], rows_per_page,
                     tuple((row[0], row[1], row[2]) for row in rows))
        previous = cache.get(key) if settings["mode"].startswith("刷新") else None
        if settings["mode"].startswith("刷新") and previous is None:
            return
        if previous is not None:
            old_signature, pages = previous
            if old_signature != signature:
                raise RuntimeError("Tool rows or page layout changed.")
            page_count = len(pages)
            for start, (sheet, table) in enumerate(pages):
                sheet.Open()
                refresh_origin_notes(part, sheet, settings, listing)
                page = rows[start * rows_per_page:(start + 1) * rows_per_page]
                table_rows = 1 + len(page) + (1 if start == len(pages) - 1 and total is not None else 0)
                refresh_production_setup_note(part, sheet, settings, table_rows, listing)
                refresh_columns(uf, table, page, total if start == len(pages) - 1 else None, settings)
            session.UpdateManager.DoUpdate(mark)
            pages[0][0].Open()
            part.DrawingSheets.RefreshCurrentSheet()
            successful = True
            return

        size = workpiece_size(part, listing)
        source_path = part.FullPath or "未保存"
        FONT_INDEX = resolve_sheet_font(part, settings["font_family"])
        pages = []
        first = None
        for page_index, start in enumerate(range(0, len(rows), rows_per_page)):
            page = rows[start:start + rows_per_page]
            name = "CAM_V31_OPTIMIZED_{}_{:02d}".format(stamp, start // rows_per_page + 1)
            sheet = part.DrawingSheets.InsertSheet(
                name, drawing.Unit.Millimeters, 297.0, 210.0)
            if first is None:
                first = sheet
            part.DrawingSheets.ActiveSheet = sheet
            add_sheet_border(part, sheet)
            fill_page(uf, sheet, group_name, page, total if page_index == len(range(0, len(rows), rows_per_page)) - 1 else None,
                      settings, 1.0, size, source_path)
            pages.append((sheet, None))
        cache[key] = (signature, pages)
        if first is not None:
            first.Open()
        part.DrawingSheets.RefreshCurrentSheet()
        successful = True
    except Exception as error:
        listing.WriteLine("CAM sheet generation failed: " + str(error))
        traceback.print_exc()
        raise
    finally:
        listing.close()
        if mark is not None and successful:
            session.UpdateManager.DoUpdate(mark)


if __name__ == "__main__":
    main()
