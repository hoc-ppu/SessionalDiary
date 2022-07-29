from abc import ABC
from datetime import timedelta
from typing import cast
from typing import Iterable
from typing import Optional

from typing import Sequence

from lxml.etree import _Element
from lxml.etree import Element
from lxml.etree import SubElement

import package.utilities as utils
from package.table_sections import AnalysisTableSection
from package.table_sections import CH_DiaryDay_TableSection
from package.table_sections import SectionParent

# from package.table_sections import TableSection
from package.table_sections import WH_DiaryDay_TableSection
from package.utilities import AID
from package.utilities import AID5
from package.utilities import CellT
from package.utilities import format_timedelta
from package.utilities import make_id_cells
from package.utilities import NS_MAP
from package.utilities import counters

# from abc import abstractmethod


class Table(ABC):
    def __init__(self, list_of_tuples: list[tuple[str, int]]):
        """Takes a list of 2 tuples of table header and cell widths"""

        # self.sections: dict[str, TableSection] | list[TableSection] = {}

        self.cols_count: int = len(list_of_tuples)  # number of cols in table

        self.xml_element = Element(
            self.__class__.__name__,
            nsmap=NS_MAP,
            attrib={
                AID + "table": "table",
                AID + "tcols": str(len(list_of_tuples)),
                AID5 + "tablestyle": "Part1Table",
                f"{AID}trows": "1",
            },
        )

        # add heading elements to table
        for item in list_of_tuples:
            heading = SubElement(
                self.xml_element,
                "Cell",
                attrib={
                    AID + "table": "cell",
                    AID + "theader": "",
                    AID + "ccolwidth": str(item[1]),
                },
            )
            heading.text = item[0]

    # def __getitem__(self, key: str):
    #     return self.sections[key]

    # @abstractmethod
    # def start_new_section(
    #     self,
    #     key: str,
    #     title: str,
    #     excel_sheet_title: str,
    #     parent: Optional[SectionParent],
    # ):
    #     ...

    # def add_new_sections(
    #     self, sects: Sequence[tuple[str, str, str, Optional[SectionParent]]]
    # ):
    #     for sec in sects:
    #         self.start_new_section(*sec)

    def increment_rows(self, increment_by: int = 1):
        trows = self.xml_element.get(AID + "trows", default=None)
        if trows:
            new_number_str = str(int(trows) + increment_by)
            self.xml_element.set(AID + "trows", new_number_str)
        else:
            self.xml_element.set(AID + "trows", "1")

    def add_table_sub_head(
        self, heading_text: str, cellstyle: str = "", subsubhead: bool = False
    ):
        """Add a subheading row to the tables xml. This is achieved by
        adding a single cell with a colspan equal to the number of
        columns in the table and adding the heading text"""

        self.increment_rows()
        cellstyle = "SubHeading"

        if subsubhead:
            cellstyle = "SubSubHeading"

        sub_head = SubElement(
            self.xml_element,
            "Cell",
            attrib={
                AID + "table": "cell",
                AID + "ccols": f"{self.cols_count}",
                AID5 + "cellstyle": cellstyle,
            },
        )
        sub_head.text = heading_text

    def add_total_duration(self, total_duration: timedelta):
        self.increment_rows()
        total_cell = utils.Body_line_below_right_align()
        total_cell.text = "Total:"
        time_cell = utils.Body_lines()
        time_cell.text = format_timedelta(total_duration)
        self.xml_element.extend(make_id_cells([None]) + [total_cell, time_cell])

    def add_section(self, section: AnalysisTableSection):
        """Add a table section to self.xml_element"""

        if section.parent is not None:
            # if there is a parent then this section of the table should have
            # a subsubsection heading rather than a subsection heading
            self.add_table_sub_head(section.title, subsubhead=True)
        else:
            self.add_table_sub_head(section.title)

        self.xml_element.extend(section.cells)
        self.increment_rows(increment_by=section.rows)
        self.add_total_duration(section.duration)

        if section.parent is not None:
            section.parent.total_duration += section.duration

        # if section.excel_sheet:
        #     # not sure that this should be here
        #     section.add_to_excel()


class Contents_Table(Table):
    def add_row(self, cells_items: Iterable[CellT]):
        self.increment_rows()
        cells = make_id_cells(cells_items)
        self.xml_element.extend(cells)


class Analysis_Table(Table):
    """class is based on etree.ElementBase
    name of the xml element will default to the name of the class
    class can be instantiated despite what type checkers may think"""

    def __init__(
        self,
        list_of_tuples: list[tuple[str, int]],
    ):
        """Takes a list of 2 tuples of table header and cell widths,
        and a dictionary of key: str, value: TableSection."""

        # self.sections_dict = sections

        self.sections: dict[str, AnalysisTableSection] = {}

        super().__init__(list_of_tuples)

    def start_new_section(
        self,
        key: str,
        title: str,
        excel_sheet_title: str,
        parent: Optional[SectionParent],
    ):
        self.sections[key] = AnalysisTableSection(title, excel_sheet_title, parent)

    def add_new_sections(
        self, sects: Sequence[tuple[str, str, str, Optional[SectionParent]]]
    ):
        for sec in sects:
            self.start_new_section(*sec)

    def __getitem__(self, key: str):
        # at the moment it looks like we need this for typing...
        return self.sections[key]


class WH_Diary_Table(Table):
    def __init__(
        self,
        list_of_tuples: list[tuple[str, int]],
    ):
        super().__init__(list_of_tuples)
        self.sections: list[WH_DiaryDay_TableSection] = []

        self.session_total_time = timedelta(seconds=0)

    def start_new_section(
        self,
        title: str,
    ):
        new_section = WH_DiaryDay_TableSection(title)
        self.sections.append(new_section)

        # also add the session total duration (so far) to the previous
        # section and add the previous section to the XML
        if len(self.sections) > 0:
            self.add_section(self.sections[-1])

        return new_section

    # ignoring The Liskov Substitution Principle here as I can't think
    # how else I would do this.
    def add_total_duration(
        self,
        daily_total_duration: timedelta,
        session_total_duration: timedelta,
    ):

        self.increment_rows()
        total_cell = utils.Right_align_cell()
        total_cell.set(AID + "ccols", "2")  # span 2 cols
        total_cell.text = "Daily Totals:"

        time_cell = utils.Body_line_above()
        time_cell.text = format_timedelta(daily_total_duration)

        self.xml_element.extend([total_cell, time_cell])

        self.increment_rows()
        total_cell = utils.Body_line_below_right_align()
        total_cell.set(AID + "ccols", "2")  # span 2 cols
        total_cell.text = "Totals for Session:"
        # print(etree.tostring(total_cell))

        time_cell = utils.Body_line_below()
        time_cell.text = format_timedelta(session_total_duration)
        # print(etree.tostring(time_cell), '\n')

        self.xml_element.extend([total_cell, time_cell])

    def add_table_sub_head(self, heading_text: str):
        super().add_table_sub_head(heading_text, "SubHeading No Toc")

    def add_section(self, section: WH_DiaryDay_TableSection):
        """Add a table section to self.xml_element"""
        self.add_table_sub_head(section.title)
        self.xml_element.extend(section.cells)
        self.increment_rows(increment_by=section.rows)
        self.add_total_duration(section.duration, self.session_total_time)


class CH_Diary_Table(Table):
    def __init__(
        self,
        list_of_tuples: list[tuple[str, int]],
    ):
        super().__init__(list_of_tuples)
        self.sections: list[CH_DiaryDay_TableSection] = []

        self.session_total_time = timedelta(seconds=0)
        self.session_aat_total = timedelta(seconds=0)

    def start_new_section(
        self,
        title: str,
    ):

        # also add the session total duration (so far) to the previous
        # section and add the previous section to the XML
        if len(self.sections) > 0:
            self.add_section(self.sections[-1])

        new_section = CH_DiaryDay_TableSection(title)
        self.sections.append(new_section)

        return new_section

    def add_total_duration(
        self,
        daily_total_duration: timedelta,
        daily_aat_total: timedelta,
        session_total_duration: timedelta,
        session_aat_total: timedelta,
    ) -> None:

        self.increment_rows()
        total_cell = utils.Right_align_cell()
        # total_cell.set(AID + 'ccols', '2')  # span 2 cols
        total_cell.text = "Daily Totals:"

        time_cell = utils.Body_line_above()
        time_cell.text = format_timedelta(daily_total_duration)

        time_2_cell = utils.Body_line_above()
        time_2_cell.text = format_timedelta(daily_aat_total)

        self.xml_element.extend(
            make_id_cells([None]) + [total_cell, time_cell, time_2_cell]
        )

        self.increment_rows()
        total_cell = utils.Body_line_below_right_align()
        # total_cell.set(AID + 'ccols', '2')  # span 2 cols
        total_cell.text = "Totals for Session:"

        time_cell = utils.Body_line_below()
        time_cell.text = format_timedelta(session_total_duration)
        time_2_cell = utils.Body_line_below()
        time_2_cell.text = format_timedelta(session_aat_total)

        cells = make_id_cells([None], attrib={AID5 + "cellstyle": "BodyLineBelow"}) + [
            total_cell,
            time_cell,
            time_2_cell,
        ]
        self.xml_element.extend(cells)

    def add_table_sub_head(self, heading_text: str):
        super().add_table_sub_head(heading_text, "SubHeading No Toc")

    def add_section(self, section: CH_DiaryDay_TableSection):
        """Add a table section to self.xml_element"""
        self.add_table_sub_head(section.title)
        # if counters.diary_cells < 3:
        #     print(f"{section.cells=}")
        #     counters.diary_cells += 1
        self.xml_element.extend(section.cells)
        self.increment_rows(increment_by=section.rows)
        self.add_total_duration(
            section.duration,
            section.after_appointed_time,
            self.session_total_time,
            self.session_aat_total,
        )


def create_contents_xml_element(
    table: Analysis_Table,
) -> _Element:

    # create XML element for the contents table
    contents_table: Contents_Table = Contents_Table(
        [
            ("Part ", 50),
            ("Contents", 200),
            ("Duration", 45),
            ("After appointed time", 45),
        ],
    )

    previous_parents: set[SectionParent] = set()
    if not isinstance(table.sections, dict):  # type: ignore
        raise TypeError("Table.sections must either be a list type or dict type")
    else:
        sections = table.sections.values()

    for table_section in sections:
        parent = table_section.parent
        if parent is not None and parent not in previous_parents:
            previous_parents.add(parent)

            table_num_dur_formatted = format_timedelta(parent.total_duration)
            try:
                table_num_aat_formatted = format_timedelta(parent.total_aat)
            except AttributeError:
                table_num_aat_formatted = ""

            table_num: str = ""
            title: str = ""
            try:
                table_num, title = parent.title.split("\t")
            except ValueError:
                title = parent.title

            cells = make_id_cells(
                [
                    f"{table_num}",
                    title,
                    f"{table_num_dur_formatted}",
                    f"{table_num_aat_formatted}",
                ],
                attrib={AID5 + "cellstyle": "RightAlign"},
            )
            contents_table.add_row(cells)  # type: ignore

            # we do not want it include tables totals more than once
            try:
                if int(table_num) == int(table_section.title.split(":\t")[0]):
                    continue
            except Exception:
                pass

        try:
            title_num, title = table_section.title.split(":\t")
        except ValueError:
            title_num = ""
            title = table_section.title
        formatted_dur = format_timedelta(table_section.duration)

        formatted_aat = ""
        # warn if aa but not enough columns for it
        if type(table_section) == AnalysisTableSection:
            table_section = cast(AnalysisTableSection, table_section)  # type: ignore
            formatted_aat: str = format_timedelta(table_section.after_appointed_time)

        cells = make_id_cells(
            [f"{title_num}", title, f"{formatted_dur}", f"{formatted_aat}"],
            attrib={AID5 + "cellstyle": "RightAlign"},
        )
        contents_table.add_row(cells)

    return contents_table.xml_element
