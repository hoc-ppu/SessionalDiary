from datetime import timedelta
from typing import Iterable, Optional, cast

from openpyxl.cell import Cell
from openpyxl.styles import Font
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from lxml import etree
from lxml.etree import SubElement

import package.utilities as utils
from package.utilities import AID, AID5
from package.utilities import make_id_cells
from package.utilities import format_timedelta

BOLD = Font(bold=True)

# exporting to excel is optional
class Excel:
    # the workbook class to output to
    out_wb: Optional[Workbook] = None


class _TableSection():
    def __init__(self, title: str):
        self.title = title
        self.rows = 0
        self.cells = []

    def __len__(self):
        return len(self.cells)

    def add_row(self, cells_items: Iterable):
        self.rows += 1
        cells = make_id_cells(cells_items)
        self.cells.extend(cells)

    def add_to(self, table):
        # ID XML stuff
        table.add_table_sub_head(self.title)
        table.extend(self.cells)
        table.increment_rows(increment_by=self.rows)


class WH_AnalysisTableSection(_TableSection):

    # create some text for the totals on the contents page
    # contents_text = ''
    # contents_dict = {}
    part_dur = timedelta(seconds=0)
    table_num_dur = {}

    # @classmethod
    # def output_contents(cls, file_path: str):
    #     with open(file_path, 'w') as f:
    #         f.write(cls.contents_text)

    def __init__(self, title: str, excel_sheet_title: str, table_num: int):
        super().__init__(title)
        self.table_num = table_num
        self.duration = timedelta(seconds=0)
        # self.out_wb = excel_workbook
        # also create an excel sheet
        # self.sheet_title = excel_sheet_title
        self.excel_sheet: Optional[Worksheet] = None
        if Excel.out_wb:
            self.excel_sheet = cast(Worksheet, Excel.out_wb.create_sheet(excel_sheet_title))

    def add_row(self, cells_items: Iterable, duration: timedelta):
        super().add_row(cells_items)
        self.duration += duration
        e_row = []
        for cell in cells_items:
            if isinstance(cell, str):
                cell = cell.replace('\t', ' ')
                e_row.append(cell)
            else:
                e_row.append(cell)
        # print(e_row)
        if self.excel_sheet:
            self.excel_sheet.append(e_row)

    def add_to(self, table):
        super().add_to(table)
        table.add_total_duration(self.duration)
        # WH_AnalysisTableSection.contents_text += (f'{self.title}\t'
        #                                           f'{format_timedelta(self.duration)}\n')
        # WH_AnalysisTableSection.contents_dict[self.title] = self.duration

        # excel stuff
        if self.excel_sheet:
            sess_tot_cell = Cell(self.excel_sheet, value='Sessional Total')
            sess_tot_cell.font = BOLD
            tot_dur_cell = Cell(self.excel_sheet, value=self.duration)  # type: ignore
            tot_dur_cell.font = BOLD

            totals_row = [None, sess_tot_cell, tot_dur_cell]

            self._add_to_excel(totals_row)

    def _add_to_excel(self, totals_row: list[Optional[Cell]]) -> None:
        if self.excel_sheet:
            self.excel_sheet.insert_rows(1, 2)
            self.excel_sheet['A1'] = self.title.replace('\t', ' ')
            # make first row bold
            self.excel_sheet['A1'].font = BOLD  # type: ignore
            for cell in self.excel_sheet[2]:  # get second row
                cell.font = BOLD  # make second row bold
            self.excel_sheet['A2'] = 'Date'
            self.excel_sheet['B2'] = 'Content'
            self.excel_sheet['C2'] = 'Duration'

            self.excel_sheet.append(totals_row)

            # tidy up the col widths of the first two columns.
            # Otherwise it's too narrow and you have to change it every time you open the excel
            self.excel_sheet.column_dimensions['A'].width = 20
            self.excel_sheet.column_dimensions['B'].width = 30


class CH_AnalysisTableSection(WH_AnalysisTableSection):

    # create some text for the totals on the contents page
    # contents_text = ''
    part_aat = timedelta(seconds=0)
    table_num_aat = {}

    def __init__(self, title: str, excel_sheet_title: str,
                 table_num: int):
        super().__init__(title, excel_sheet_title, table_num)
        self.after_appointed_time = timedelta(seconds=0)

    def add_row(self, cells_items: Iterable, duration: timedelta, aat: timedelta):
        super().add_row(cells_items, duration=duration)
        self.after_appointed_time += aat

    def add_to(self, table):
        table.add_table_sub_head(self.title)
        table.extend(self.cells)
        table.increment_rows(increment_by=self.rows)
        table.add_total_duration(self.duration, self.after_appointed_time)
        # CH_AnalysisTableSection.contents_text += (f'{self.title}\t'
        #                                   f'{format_timedelta(self.duration)}\t'
        #                                   f'{format_timedelta(self.after_appointed_time)}\n')
        CH_AnalysisTableSection.part_dur += self.duration
        # part_dur = format_timedelta(CH_AnalysisTableSection.part_dur)
        # self_dur = format_timedelta(self.duration)
        # print(f'{self.table_num}:\t{part_dur=}, {self_dur=}')
        CH_AnalysisTableSection.part_aat += self.after_appointed_time

        if self.table_num in CH_AnalysisTableSection.table_num_dur:
            CH_AnalysisTableSection.table_num_dur[self.table_num] += self.duration
        else:
            CH_AnalysisTableSection.table_num_dur[self.table_num] = self.duration

        if self.table_num in CH_AnalysisTableSection.table_num_aat:
            CH_AnalysisTableSection.table_num_aat[self.table_num] += self.after_appointed_time
        else:
            CH_AnalysisTableSection.table_num_aat[self.table_num] = self.after_appointed_time

        # excel stuff
        if self.excel_sheet:
            sess_tot_cell = Cell(self.excel_sheet, value='Sessional Total')
            sess_tot_cell.font = BOLD
            tot_dur_cell = Cell(self.excel_sheet, value=self.duration)  # type: ignore
            tot_dur_cell.font = BOLD
            tot_aat_cell = Cell(self.excel_sheet, value=self.after_appointed_time)  # type: ignore
            tot_aat_cell.font = BOLD

            totals_row = [None, sess_tot_cell, tot_dur_cell, tot_aat_cell]

            self._add_to_excel(totals_row)

        # cell_values = [cell.text for cell in self.cells]
        # self.excel_sheet.append(cell_values)
        # self.excel_sheet.insert_rows(1, 2)
        # self.excel_sheet['A1'] = self.title.replace('\t', ' ')
        # self.excel_sheet['A1'].font = BOLD  # make first row bold
        # for cell in self.excel_sheet[2]:  # get second row
        #     cell.font = BOLD  # make second row bold
        # self.excel_sheet['A2'] = 'Date'
        # self.excel_sheet['B2'] = 'Content'
        # self.excel_sheet['C2'] = 'Duration'
        # self.excel_sheet['D2'] = 'After appointed time'

        # sess_tot_cell = Cell(self.excel_sheet, value='Sessional Total')
        # sess_tot_cell.font = BOLD
        # tot_dur_cell = Cell(self.excel_sheet, value=self.duration)  # type: ignore
        # tot_dur_cell.font = BOLD
        # tot_aat_cell = Cell(self.excel_sheet, value=self.after_appointed_time)  # type: ignore
        # tot_aat_cell.font = BOLD

        # totals_row = [None, sess_tot_cell, tot_dur_cell, tot_aat_cell]
        # self.excel_sheet.append(totals_row)

        # # tidy up the col widths of the first two columns.
        # # Otherwise it's too narrow and you have to change it every time you open the excel
        # self.excel_sheet.column_dimensions['A'].width = 20
        # self.excel_sheet.column_dimensions['B'].width = 30

    def _add_to_excel(self, totals_row: list[Optional[Cell]]):
        if self.excel_sheet:
            super()._add_to_excel(totals_row)
            # the chamber is different from WH as it includes an extra col
            self.excel_sheet['D2'] = 'After appointed time'  # extra col head


class WH_DiaryDay_TableSection(_TableSection):

    def __init__(self, title: str):
        super().__init__(title)
        self.duration = timedelta(seconds=0)

    def add_row(self, cells, duration: timedelta):
        super().add_row(cells)
        self.duration += duration


    def add_to(self, table, session_total_time: timedelta):
        table.add_table_sub_head(self.title)
        table.extend(self.cells)
        table.increment_rows(increment_by=self.rows)
        table.add_total_duration(self.duration, session_total_time)


class CH_DiaryDay_TableSection(WH_DiaryDay_TableSection):

    def __init__(self, title: str):
        super().__init__(title)
        self.after_appointed_time = timedelta(seconds=0)

    def add_row(self, cells,
                duration: timedelta, aat: timedelta):
        super().add_row(cells, duration)
        self.after_appointed_time += aat

    def add_to(self, table,
               session_duration: timedelta, session_aat: timedelta):
        table.add_table_sub_head(self.title)
        table.extend(self.cells)
        table.increment_rows(increment_by=self.rows)
        table.add_total_duration(self.duration,
                                 self.after_appointed_time,
                                 session_duration,
                                 session_aat)


class WH_Table(etree.ElementBase):
    """class is based on etree.ElementBase
    name of the xml element will default to the name of the class
    class can be instantiated despite what type checkers may think"""

    def increment_rows(self, increment_by=1):
        trows = self.get(AID + 'trows', default=None)
        if trows:
            self.set(AID + 'trows', str(int(trows) + increment_by))
        else:
            self.set(AID + 'trows', '1')

    def add_total_duration(self, total_duration: timedelta):
        self.increment_rows()
        total_cell = utils.Body_line_below_right_align()
        # total_cell.set(AID + 'ccols', '2')  # span 2 cols
        # total_cell.set(AID5 + 'cellstyle', 'BodyLineBelowRightAlign')  # right align content
        total_cell.text = 'Total:'
        time_cell = utils.Body_lines()
        # time_cell.set(AID5 + 'cellstyle', 'BodyLines')
        time_cell.text = format_timedelta(total_duration)
        self.extend(make_id_cells([None]) + [total_cell, time_cell])  # type: ignore

    def add_table_sub_head(self, heading_text: str):
        self.increment_rows()
        sub_head = SubElement(self, 'Cell',
                              attrib={AID + 'table': 'cell',
                                      AID + 'ccols': '3',
                                      AID5 + 'cellstyle': 'SubHeading'})
        sub_head.text = heading_text


class WH_Diary_Table(WH_Table):

    def add_total_duration(self, daily_total_duration, session_total_duration):
        self.increment_rows()
        total_cell = utils.Right_align_cell()
        total_cell.set(AID + 'ccols', '2')  # span 2 cols
        total_cell.text = 'Daily Totals:'

        time_cell = utils.Body_line_above()
        time_cell.text = format_timedelta(daily_total_duration)

        self.extend([total_cell, time_cell])  # type: ignore

        self.increment_rows()
        total_cell = utils.Body_line_below_right_align()
        total_cell.set(AID + 'ccols', '2')  # span 2 cols
        total_cell.text = 'Totals for Session:'

        time_cell = utils.Body_line_below()
        time_cell.text = format_timedelta(session_total_duration)

        self.extend([total_cell, time_cell])  # type: ignore

    def add_table_sub_head(self, heading_text: str):
        self.increment_rows()
        sub_head = SubElement(self, 'Cell',
                              attrib={AID + 'table': 'cell',
                                      AID + 'ccols': '3',
                                      AID5 + 'cellstyle': 'SubHeading No Toc'})
        sub_head.text = heading_text


class CH_Table(WH_Table):

    def add_total_duration(self, total_duration, aat_total) -> None:
        self.increment_rows()
        total_cell = utils.Body_line_below_right_align()
        total_cell.text = 'Total:'

        time_cell = utils.Body_lines()
        time_cell.text = format_timedelta(total_duration)

        time_2_cell = utils.Body_lines()
        time_2_cell.text = format_timedelta(aat_total)

        self.extend(make_id_cells([None]) + [total_cell, time_cell, time_2_cell])  # type: ignore

    def add_table_sub_head(self, heading_text: str):
        self.increment_rows()
        SubElement(self, 'Cell',
                   attrib={AID + 'table': 'cell',
                           AID + 'ccols': '4',
                           AID5 + 'cellstyle': 'SubHeading'}
                   ).text = heading_text


class CH_Diary_Table(WH_Table):

    def add_total_duration(self, daily_total_duration, daily_aat_total,
                           session_total_duration, session_aat_total) -> None:
        self.increment_rows()
        total_cell = utils.Right_align_cell()
        total_cell.set(AID + 'ccols', '2')  # span 2 cols
        total_cell.text = 'Daily Totals:'

        time_cell = utils.Body_line_above()
        time_cell.text = format_timedelta(daily_total_duration)

        time_2_cell = utils.Body_line_above()
        time_2_cell.text = format_timedelta(daily_aat_total)

        self.extend(make_id_cells([None]) + [total_cell, time_cell, time_2_cell])  # type: ignore


        self.increment_rows()
        total_cell = utils.Body_line_below_right_align()
        total_cell.set(AID + 'ccols', '2')  # span 2 cols
        total_cell.text = 'Totals for Session:'

        time_cell = utils.Body_line_below()
        time_cell.text = format_timedelta(session_total_duration)
        time_2_cell = utils.Body_line_below()
        time_2_cell.text = format_timedelta(session_aat_total)

        cells = (make_id_cells([None], attrib={AID5 + 'cellstyle': 'BodyLineBelow'})
                 + [total_cell, time_cell, time_2_cell])
        self.extend(cells)   # type: ignore

    def add_table_sub_head(self, heading_text: str):
        self.increment_rows()
        SubElement(self, 'Cell',
                   attrib={AID + 'table': 'cell',
                           AID + 'ccols': '5',
                           AID5 + 'cellstyle': 'SubHeading No Toc'}
                   ).text = heading_text
