from datetime import timedelta, datetime, date
from copy import deepcopy
from pathlib import Path
from typing import Iterable
from typing import Union
from typing import Any
from typing import cast
from typing import Optional

from lxml import etree
from lxml.etree import Element
from lxml.etree import _Element
from lxml.etree import iselement
from openpyxl import Workbook
from openpyxl.styles import Font

# a cell element (for InDesign) or its contents
CellT = Union[str, _Element, timedelta, None, int, float]

AID = "{http://ns.adobe.com/AdobeInDesign/4.0/}"
AID5 = "{http://ns.adobe.com/AdobeInDesign/5.0/}"

NS_MAP: dict[str | None, str] = {
    "aid": "http://ns.adobe.com/AdobeInDesign/4.0/",
    "aid5": "http://ns.adobe.com/AdobeInDesign/5.0/",
}

BOLD = Font(bold=True)


class debug:
    debug = False


class counters:
    diary_add_row = 0
    diary_cells = 0
    tables_abc_add_row = 0
    section = 0


# exporting to excel is optional
class Excel:
    # the workbook class to output to
    out_wb: Optional[Workbook] = None


# template for an InDesign table cell
id_cell = Element("Cell", attrib={AID + "table": "cell"})


def ID_Cell() -> _Element:
    """Create a XML cell Element for InDesign"""
    return deepcopy(id_cell)


def Right_align_cell() -> _Element:
    """Create a XML cell Element with the RightAlign cellstyle applied."""

    cell = ID_Cell()
    cell.set(AID5 + "cellstyle", "RightAlign")
    return cell


def Body_line_below_right_align() -> _Element:
    """Create a XML cell Element with the
    BodyLineBelowRightAlign cellstyle applied."""

    cell = ID_Cell()
    cell.set(AID5 + "cellstyle", "BodyLineBelowRightAlign")
    return cell


def Body_line_below() -> _Element:
    """Create a XML cell Element with the
    BodyLineBelow cellstyle applied."""

    cell = ID_Cell()
    cell.set(AID5 + "cellstyle", "BodyLineBelow")
    return cell


def Body_line_above() -> _Element:
    """Create a XML cell Element with the
    BodyLineAbove cellstyle applied."""

    cell = ID_Cell()
    cell.set(AID5 + "cellstyle", "BodyLineAbove")
    return cell


def Body_lines() -> _Element:
    """Create a XML cell Element with the
    BodyLines cellstyle applied."""

    cell = ID_Cell()
    cell.set(AID5 + "cellstyle", "BodyLines")
    return cell


def make_id_cells(
    iterable: Iterable[CellT], attrib: dict[str, str] = {}
) -> list[_Element]:
    cells: list[_Element] = []
    for item in iterable:
        if iselement(item):
            # type checker doesn't know about iselement it seems
            item = cast(_Element, item)
            cells.append(deepcopy(item))
        else:
            cell = ID_Cell()

            for attribute_key, attribute_value in attrib.items():
                # there may not be any attributes of corse
                cell.set(attribute_key, attribute_value)
            if isinstance(item, str):
                cell.text = item
            elif isinstance(item, timedelta):
                cell.text = format_timedelta(item)
            elif item is None:
                cell.text = ""
            else:
                cell.text = str(item)
            cells.append(cell)
    return cells


def format_timedelta(td: timedelta) -> str:
    total_seconds = td.total_seconds()
    hours = round(total_seconds // 3600)
    mins = round(total_seconds % 3600 / 60)
    return f"{hours}.{mins:02}"


def format_date(date_containing_item: Union[datetime, date, str]):
    if isinstance(date_containing_item, datetime):
        return date_containing_item.strftime("%a,\t%d\t%b\t%Y")
    if isinstance(date_containing_item, str):
        try:
            return datetime.strptime(date_containing_item, "%d %B %Y").strftime(
                "%a,\t%d\t%b\t%Y"
            )
        except ValueError:
            # TODO: log this
            print("print")
            print(date_containing_item)


# def timedelta_from_time(t: time, default=timedelta(seconds=0)):
#     try:
#         return datetime.combine(date.min, t) - datetime.min
#     except TypeError as e:
#         if t != '':
#             print(f'{t=}')
#             print(e)
#         return default


def str_strip(input: Any) -> str:
    """Return empty string if input is none otherwise return str(input).strip()"""

    if input is None:
        output = ""
    else:
        output = str(input).strip()

    return output


def write_xml(lxml_element: _Element, path: Path):
    output_root = Element("root")
    output_root.append(lxml_element)
    output_tree = etree.ElementTree(output_root)
    output_tree.write(
        str(path),
        encoding="UTF-8",
        xml_declaration=True,
    )

    if debug.debug:
        # if in debug mode also export an indented version
        # as it's easier to do compares on.
        etree.indent(output_tree, space="  ")
        output_tree.write(
            str(path.with_name(f"indented_{path.name}")),
            encoding="UTF-8",
            xml_declaration=True,
        )
    print(f"Created: {path}")
