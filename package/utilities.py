from datetime import timedelta, datetime, date, time
from copy import deepcopy
from typing import Iterable
from typing import Union
from typing import Any


from lxml.etree import Element
from lxml.etree import _Element
from lxml.etree import iselement

AID  = '{http://ns.adobe.com/AdobeInDesign/4.0/}'
AID5 = '{http://ns.adobe.com/AdobeInDesign/5.0/}'

NS_MAP = {'aid':  'http://ns.adobe.com/AdobeInDesign/4.0/',
          'aid5': 'http://ns.adobe.com/AdobeInDesign/5.0/'}

# template for an InDesign table cell
id_cell = Element('Cell', attrib={AID + 'table': 'cell'})

def ID_Cell() -> _Element:
    """Create a XML cell Element for InDesign"""
    return deepcopy(id_cell)


def Right_align_cell() -> _Element:
    """Create a XML cell Element with the RightAlign cellstyle applied."""

    cell = ID_Cell()
    cell.set(AID5 + 'cellstyle', 'RightAlign')
    return cell


def Body_line_below_right_align() -> _Element:
    """Create a XML cell Element with the
    BodyLineBelowRightAlign cellstyle applied."""

    cell = ID_Cell()
    cell.set(AID5 + 'cellstyle', 'BodyLineBelowRightAlign')
    return cell


def Body_line_below() -> _Element:
    """Create a XML cell Element with the
    BodyLineBelow cellstyle applied."""

    cell = ID_Cell()
    cell.set(AID5 + 'cellstyle', 'BodyLineBelow')
    return cell


def Body_line_above() -> _Element:
    """Create a XML cell Element with the
    BodyLineAbove cellstyle applied."""

    cell = ID_Cell()
    cell.set(AID5 + 'cellstyle', 'BodyLineAbove')
    return cell


def Body_lines() -> _Element:
    """Create a XML cell Element with the
    BodyLines cellstyle applied."""

    cell = ID_Cell()
    cell.set(AID5 + 'cellstyle', 'BodyLines')
    return cell


def make_id_cells(iterable: Iterable[Union[str, _Element, timedelta, None, int, float]],
                  attrib: dict = {}) -> list[_Element]:
    cells = []
    for item in iterable:
        if iselement(item):
            cells.append(deepcopy(item))
        else:
            cell = ID_Cell()
            if attrib:
                for attribute_key, attribute_value in attrib.items():
                    cell.set(attribute_key, attribute_value)
            if isinstance(item, str):
                cell.text = item
            elif isinstance(item, timedelta):
                cell.text = format_timedelta(item)
            elif item is None:
                cell.text = ''
            else:
                cell.text = str(item)
            cells.append(cell)
    return cells


def format_timedelta(td: timedelta) -> str:
    total_seconds = td.total_seconds()
    hours = round(total_seconds // 3600)
    mins = round(total_seconds % 3600 / 60)
    return f'{hours}.{mins:02}'


def format_date(date_containing_item: Union[datetime, date, str]):
    if isinstance(date_containing_item, datetime):
        return date_containing_item.strftime('%a,\t%d\t%b\t%Y')
    if isinstance(date_containing_item, str):
        try:
            return datetime.strptime(date_containing_item, '%d %B %Y').strftime('%a,\t%d\t%b\t%Y')
        except ValueError:
            # TODO: log this
            print('print')
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
    '''Return empty string if input is none otherwise return str(input)'''

    if input is None:
        output = ''
    else:
        output = str(input).strip()

    return output
