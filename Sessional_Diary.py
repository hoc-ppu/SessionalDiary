#!/usr/bin/env python3

import argparse
from datetime import date, datetime, timedelta, time
import os
# from pathlib import Path
import sys
# import tkinter as tk
# from tkinter import ttk, filedialog, messagebox
from typing import cast
from typing import Type
from typing import Sequence

# 3rd party imports
from lxml import etree
from lxml.etree import Element
from lxml.etree import SubElement
from openpyxl import load_workbook, Workbook
from openpyxl.styles.numbers import STRIP_RE
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell import cell as CELL
from openpyxl.cell.cell import Cell

# 1st party imports
from package.utilities import ID_Cell
from package.tables import CH_Table, WH_Table
from package.tables import CH_Diary_Table, WH_Diary_Table
from package.tables import CH_DiaryDay_TableSection, WH_DiaryDay_TableSection
from package.tables import CH_AnalysisTableSection, WH_AnalysisTableSection
from package.tables import Excel
from package.utilities import make_id_cells, format_timedelta
from package.utilities import str_strip
from package.utilities import timedelta_from_time, format_date
from package.utilities import AID, AID5, NS_MAP


# override default openpyxl timedelta (duration) format
CELL.TIME_FORMATS[timedelta] = '[h].mm'


DATE_NUM_LOOK_UP: dict[date, int] = {}


#  We expect the following column headings for the
# Chamber section in the Excel document
DAY = 'Day'
DATE = 'Date'
TIME = 'Time'
SUBJECT1 = 'Subject 1'
SUBJECT2 = 'Subject 2'
TAGS = 'Tags'
DURATION = 'DurationFx'
AAT = 'AAT'

CHAMBER_COLS = [DAY, DATE, TIME, SUBJECT1, SUBJECT2, TAGS, DURATION, AAT]
WH_COLS = [DAY, DATE, TIME, SUBJECT1, SUBJECT2, TAGS, DURATION]

CH_SHEET_TITLE = 'Chamber'
WH_SHEET_TITLE = 'Westminster Hall'

class WHRow:
    title_index: dict[str, int] = {}

    def __init__(self, excel_row: Sequence[Cell]):
        t_index = WHRow.title_index

        self.inner_init(excel_row, t_index)

    def inner_init(self, excel_row: Sequence[Cell], t_index: dict[str, int]):

        if not t_index:
            print('Error: title_index not set up')
            exit()


        self.day: int
        _day = excel_row[t_index[DAY]].value
        if isinstance(_day, int):
            self.day = _day
        else:
            print(excel_row[t_index[DAY]].coordinate,
                  ' has value ',
                  _day)
            raise ValueError

        _date = excel_row[t_index[DATE]].value

        self.date: date
        if isinstance(_date, date):
            self.date = _date
        else:
            print(excel_row[t_index[DATE]].coordinate,
                  ' has value ',
                  _date)
            raise ValueError

        self.time: time
        time_cell = excel_row[t_index[TIME]]
        _time = time_cell.value
        if isinstance(_time, time):
            self.time = _time
        else:
            print(excel_row[t_index[TIME]].coordinate,
                  ' has value ',
                  _date)
            raise ValueError

        if isinstance(self.time, datetime):
            time_obj = self.time.time()
            print(f'There is a datetime at cell {time_cell.coordinate}:', str(self.time))
            print(f'This has been converted to the following time: {time_obj}')
            self.time = time_obj

        self.subject1: str = str_strip(
            excel_row[t_index[SUBJECT1]].value)

        self.subject2: str = str_strip(
            excel_row[t_index[SUBJECT2]].value)

        self.tags = str_strip(
            excel_row[t_index[TAGS]].value)


        duration_cell = excel_row[t_index[DURATION]]
        self.duration: timedelta
        if isinstance(duration_cell.value, datetime):
            # don't trust the datetime only the time
            time_obj = duration_cell.value.time()
            print(f'There is a datetime at cell {duration_cell.coordinate}:',
                  str(duration_cell.value))
            print(f'This has been converted to the following time: {time_obj}')
            self.duration = datetime.combine(date.min, time_obj) - datetime.min
        elif isinstance(duration_cell.value, time):
            self.duration = datetime.combine(date.min, duration_cell.value) - datetime.min
        elif not isinstance(duration_cell.value, timedelta):
            # log this
            # print(f'Problem in cell {duration_cell.coordinate}')
            self.duration = timedelta()

class CHRow(WHRow):
    title_index: dict[str, int] = {}

    def __init__(self, excel_row: Sequence[Cell]):

        t_index = CHRow.title_index
        super().inner_init(excel_row, t_index)

        aat_cell = excel_row[t_index[AAT]]
        self.aat: timedelta
        if isinstance(aat_cell.value, datetime):
            # don't trust the datetime only the time
            time_obj = aat_cell.value.time()
            print(f'There is a datetime at cell {aat_cell.coordinate}:', str(aat_cell.value))
            print(f'This has been converted to the following time: {time_obj}')
            self.aat = datetime.combine(date.min, time_obj) - datetime.min
        elif isinstance(aat_cell.value, time):
            self.aat = datetime.combine(date.min, aat_cell.value) - datetime.min
        elif not isinstance(aat_cell.value, timedelta):
            self.aat = timedelta()




class Sessional_Diary:

    # expected column headings for the chamber
    # the order does not matter

    def __init__(self, input_excel_file_path: str, no_excel: bool):
        self.input_workbook = load_workbook(filename=input_excel_file_path,
                                            data_only=True, read_only=True)

        # if we require an output excel file
        if no_excel is False:
            Excel.out_wb = Workbook()  # new Excel workbook obi


    def check_chamber(self):
        try:
            cmbr_data = cast(Worksheet, self.input_workbook[CH_SHEET_TITLE])
        except Exception:
            print('There is no "Chamber" worksheet in the Excel file.',
                  'This sheet is required.')
            exit()

        top_row = cmbr_data[1]
        CHRow.title_index = {item.value: i for i, item in enumerate(top_row)}

        if not set(CHAMBER_COLS).issubset(set(CHRow.title_index.keys())):
            expected_row_headings = '", "'.join(CHAMBER_COLS)
            print(f'Expected the following column titles '
                  f'to be in the top row of the {CHAMBER_COLS} sheet\n',
                  f'"{expected_row_headings}"')

    def check_wh(self):
        try:
            wh_data = cast(Worksheet, self.input_workbook[WH_SHEET_TITLE])
        except Exception:
            print(f'There is no "{WH_SHEET_TITLE}" worksheet in the Excel file.',
                  'This sheet is required.')
            exit()

        top_row = wh_data[1]
        WHRow.title_index = {item.value: i for i, item in enumerate(top_row)}

        if not set(WH_COLS).issubset(set(WHRow.title_index.keys())):
            expected_row_headings = '", "'.join(WH_COLS)
            print(f'Expected the following column titles '
                  f'to be in the top row of the {WH_SHEET_TITLE} sheet',
                  f'"{expected_row_headings}"',
                  f'Got {WHRow.title_index.keys()}',
                  sep='\n')



    def house_diary(self, output_folder_path: str = ''):
        """Create an (indesign formatted) XML file for the house diary section of
        the Sessional diary."""

        self.check_chamber()

        cmbr_data = cast(Worksheet, self.input_workbook[CH_SHEET_TITLE])

        session_total_time      = timedelta(seconds=0)
        # day_total_time          = timedelta(seconds=0)
        session_total_after_moi = timedelta(seconds=0)
        # day_total_after_moi     = timedelta(seconds=0)

        table_sections = []

        # day_number_counter = 0

        table_ele = id_table(
            [('Time', 35), ('Subject', 310),
             ('Exit', 45), ('Duration', 45),
             ('After appointed time', 45)],
            table_class=CH_Diary_Table)

        previous_day = 1

        for c, excel_row in enumerate(cmbr_data.iter_rows()):
            if c == 0:
                continue

            # check to see if all items in list are '' as there can be blank rows
            if all(not v for v in excel_row[:10]):
                continue

            entry = CHRow(excel_row)

            if c == 1:
                table_sections.append(CH_DiaryDay_TableSection(
                    f'{entry.day}.\u2002{entry.date.strftime("%A %d %B %Y")}'))


            if entry.day != previous_day:
                previous_day = entry.day

                # day_total_time = timedelta()
                # day_total_after_moi = timedelta()

                table_sections[-1].add_to(table_ele, session_total_time, session_total_after_moi)

                table_sections.append(CH_DiaryDay_TableSection(
                    f'{entry.day}.\u2002{entry.date.strftime("%A %d %B %Y")}'))

                # add the date and number to the lookup.
                # this is so this info can also be put in the WH able
                DATE_NUM_LOOK_UP[entry.date] = entry.day

            # else:
            #     day_total_time += entry.duration
            #     day_total_after_moi += entry.aat

            # need to add up all the durations
            session_total_time += entry.duration
            session_total_after_moi += entry.aat

            # there will be 5 cells per row
            cell = ID_Cell()

            # create a Bold element. Optionally can have non bold tail text
            bold = SubElement(cell, 'Bold')
            bold.text = entry.subject1
            if entry.subject2:
                bold.tail = entry.subject2.strip()  # this text will not be bold


            duration = entry.duration
            if duration == timedelta():
                duration = ''
            aat = entry.aat
            if aat == timedelta():
                aat = ''

            table_sections[-1].add_row(
                [entry.time.strftime('%H.%M'), cell, entry.tags, duration, aat],
                duration=entry.duration, aat=entry.aat)

        # need to add the last table section
        if len(table_sections) > 0:
            table_sections[-1].add_to(table_ele, session_total_time, session_total_after_moi)


        # now output XML (for InDesign) file
        output_root = Element('root')  # create root element
        output_root.append(table_ele)
        output_tree = etree.ElementTree(output_root)
        output_tree.write(os.path.join(output_folder_path, 'House_Diary.xml'),
                          encoding='UTF-8', xml_declaration=True)

    def house_analysis(self, output_folder_path: str = ''):

        # we're only interested in the main data here
        main_data = self.input_workbook[CH_SHEET_TITLE]

        # add heading elements to table
        table_ele = id_table(
            [('Date', 95), ('', 295), ('Duration', 45), ('After appointed time', 45)],
            table_class=CH_Table
        )

        t_sections = {
            # the order matters!
            'addresses': CH_AnalysisTableSection(
                '1:\tAddresses other than Prayers',
                '1 Addresses other than Prayers',
                1),
            'second_readings': CH_AnalysisTableSection(
                '2a:\tGovernment Bills: Read a second time and committed to Public Bill Committee',
                '2a Govt Bills 2R & committed',
                2),
            'cwh_bills': CH_AnalysisTableSection(
                '2b:\tGovernment Bills: Read a second time and committed to '
                'Committee of the whole House (in whole or part)',
                '2b Govt Bill 2R & sent to CWH',
                2),
            'cwh_2_bills': CH_AnalysisTableSection(
                '2d:\tGovernment Bills: Committee of the whole House',
                '2d Govt Bills CWH',
                2),
            'gov_bil_cons': CH_AnalysisTableSection(
                '2e:\tGovernment Bills: Consideration',
                '2e Govt Bills Consideration',
                2),
            'gov_bill_3rd': CH_AnalysisTableSection(
                '2f:\tGovernment Bills: Third Reading',
                '2f Govt Bills 3R',
                2),
            'gov_bill_lord_amend': CH_AnalysisTableSection(
                '2g:\tGovernment Bills: Lord Amendments',
                '2g Lords Amendments',
                2),
            'alloc_time': CH_AnalysisTableSection(
                '2h:\tAllocation of time motions',
                '2h Allocation of time motions',
                2),
            'gov_bill_other': CH_AnalysisTableSection(
                '2i:\tGovernment Bills: Other Stages',
                '2i Govt Bills Other Stages',
                2),
            'pmbs_2r': CH_AnalysisTableSection(
                '3a:\tPrivate Members\' Bills: Second Reading',
                '3a PMB 2R',
                3),
            'pmbs_other': CH_AnalysisTableSection(
                '3b:\tPrivate Members\' Bills: Other Stages',
                '3b PMB Other stages',
                3),
            'private_business': CH_AnalysisTableSection(
                '4a:\tPrivate Business',
                '4a Private Business',
                4),
            'eu_docs': CH_AnalysisTableSection(
                '5a:\tEuropean Union documents',
                '5a European Union documents',
                5),
            'gov_motions': CH_AnalysisTableSection(
                '5b:\tGovernment motions',
                '5b Government motions',
                5),
            'gov_motions_gen': CH_AnalysisTableSection(
                '5c:\tGovernment motions (General)',
                '5c Govt motions (General)',
                5),
            'gen_debates': CH_AnalysisTableSection(
                '5d:\tGovernment motions (General Debates)',
                '5d Govt motions (Gen Debates)',
                5),
            'opposition_days': CH_AnalysisTableSection(
                '6a:\tOpposition Days',
                '6a Opposition Days',
                6),
            'oppo_motions_in_gov_time': CH_AnalysisTableSection(
                '6b:\tOpposition motions in Government time',
                '6b Opp Motion in Govt time',
                6),
            'backbench_business': CH_AnalysisTableSection(
                '7: \tBackbench Business',
                '7 Backbench Business',
                7),
            'pm_motion': CH_AnalysisTableSection(
                '8a:\tPrivate Members\' Motions',
                '8a Private Members\' Motions',
                8),
            'ten_min_motion': CH_AnalysisTableSection(
                '8b:\tTen Minute Rule Motions',
                '8b Ten minute rules',
                8),
            'emergency_debates': CH_AnalysisTableSection(
                '8c:\tEmergency debates',
                '8c Emergency debates',
                8),
            'adjournment_debates': CH_AnalysisTableSection(
                '8d:\tAdjournment debates',
                '8d Adjournment debates',
                8),
            'estimates': CH_AnalysisTableSection(
                '9:\tEstimates',
                '9 Estimates',
                9),
            'money': CH_AnalysisTableSection(
                '10:\tMoney Resolutions',
                '10 Money Resolutions',
                10),
            'ways_and_means': CH_AnalysisTableSection(
                '11:\tWays and Means',
                '11 Ways and Means',
                11),
            'affirmative_sis': CH_AnalysisTableSection(
                '12:\tAffirmative Statutory Instruments',
                '12 Affirmative SIs',
                12),
            'negative_sis': CH_AnalysisTableSection(
                '13:\tNegative Statutory Instruments',
                '13 Negative SIs',
                13),
            'questions': CH_AnalysisTableSection(
                '14a:\tQuestions',
                '14a Questions',
                14),
            'topical_questions': CH_AnalysisTableSection(
                '14b:\tTopical Questions',
                '14b Topical Questions',
                14),
            'urgent_questions': CH_AnalysisTableSection(
                '14c:\tUrgent Questions',
                '14c Urgent Questions',
                14),
            'statements': CH_AnalysisTableSection(
                '14d:\tStatements',
                '14d Statements',
                14),
            'business_statements': CH_AnalysisTableSection(
                '14e:\tBusiness Statements',
                '14e Business Statements',
                14),
            'committee_statements': CH_AnalysisTableSection(
                '14f:\tCommittee Statements',
                '14f Committee Statements',
                14),
            'app_for_emerg_debate': CH_AnalysisTableSection(
                '14g:\tS.O. No. 24 Applications',
                '14g SO No 24 Applications',
                14),
            'points_of_order': CH_AnalysisTableSection(
                '14h:\tPoints of Order',
                '14h Points of Order',
                14),
            'public_petitions': CH_AnalysisTableSection(
                '14i:\tPublic Petitions',
                '14i Public Petitions',
                14),
            'miscellaneous': CH_AnalysisTableSection(
                '14j:\tMiscellaneous',
                '14j Miscellaneous',
                14),
            'prayers': CH_AnalysisTableSection(
                '15:\tDaily Prayers',
                '15 Daily Prayers',
                15),
        }

        for c, excel_row in enumerate(main_data.iter_rows()):  # type: ignore

            if c == 0:
                # top row just has headings in
                continue

            # only interested in vfg the first few cells
            # row_values = [item.value for item in excel_row[:9]]

            # check to see if all items in list are None as there are lots of blank rows
            if all(v is None for v in excel_row[:9]):
                continue

            # sometimes the value in the cell is not a time but is instead a datetime
            # for j in (7, 8):
            #     if isinstance(row_values[j], datetime):
            #         time_obj = row_values[j].time()
            #         print(f'There is a datetime at row {c + 1}: {row_values[j]}.  '
            #               f'Converting to time: {time_obj}')
            #         row_values[j] = time_obj

            # # we want empty cells to be returned as ''
            # for i, value in enumerate(row_values):
            #     if value is None:
            #         row_values[i] = ''

            entry = CHRow(excel_row)

            # unpack the row_values into variables
            # col_day        = row_values[0]
            # col_date       = row_values[1]
            # # col_time     = row_values[2]  # not needed
            # col_subject    = row_values[3]
            # col_subject2   = row_values[4]
            # # col_subject3   = row_values[5]
            # col_exit       = row_values[6]
            # col_duration   = row_values[7]
            # col_a_a_t      = row_values[8]

            # if entry.day == 'Day' and col_date == 'Date':
            #     # ignore rows with column titles in
            #     continue

            # if not isinstance(col_date, date):
            #     # assume we can skip any records without a date
            #     continue

            forematted_date = format_date(entry.date)  # type: ignore

            # col_time = Time.strip().lower()
            subject_lower = entry.subject1.lower()
            col_exit    = entry.tags.lower()

            # we need timedelta objects
            # col_duration_td = timedelta_from_time(entry.duration, default=timedelta(seconds=0))
            # col_a_a_t_td = timedelta_from_time(entry.aat, default=timedelta(seconds=0))

            cells_vals = [
                forematted_date,
                entry.subject2,
                entry.duration,
                entry.aat
            ]

            fullrow = [cells_vals, entry.duration, entry.aat]

            # Table 1 Addresses other than Prayers
            if subject_lower == 'address':
                t_sections['addresses'].add_row(*fullrow)
            # Table 2a Government bills second reading
            if '[pmb]' not in col_exit:
                # here we have items that are not explicitly private members' bills
                if subject_lower == 'second reading' and 'pbc' in col_exit:
                    # gov bill second reading
                    t_sections['second_readings'].add_row(*fullrow)

                if 'committee of the whole house' in subject_lower:
                    t_sections['cwh_2_bills'].add_row(*fullrow)
                if 'consideration' in subject_lower:
                    # gov bill consideration
                    t_sections['gov_bil_cons'].add_row(*fullrow)
                if subject_lower == 'third reading':
                    # gov bill third reading
                    t_sections['gov_bill_3rd'].add_row(*fullrow)
                if subject_lower == 'lords amendments':
                    # gov bill lords amendments
                    t_sections['gov_bill_lord_amend'].add_row(*fullrow)
                gov_bill_other_subs = ('second and third reading',
                                       'money resolution',
                                       'lords amendments')
                if ('legislative grand committee' in subject_lower
                        or subject_lower in gov_bill_other_subs):
                    t_sections['gov_bill_other'].add_row(*fullrow)

            if (subject_lower == 'second reading'
                    and 'committee of the whole house' in entry.subject2.lower()):
                t_sections['cwh_bills'].add_row(*fullrow)

            if subject_lower.lower() in ('general debate', 'general motion'):
                t_sections['alloc_time'].add_row(*fullrow)

            if '[pmb]' in col_exit:
                if subject_lower == 'second reading':
                    # private members' bills second reading
                    t_sections['pmbs_2r'].add_row(*fullrow)
                elif subject_lower not in ('ten minute rule motion',
                                           'point of order', 'remaining orders'):
                    # private members' bills other
                    # this does not include ten minute rules
                    t_sections['pmbs_other'].add_row(*fullrow)

            if 'private business' in subject_lower:
                t_sections['private_business'].add_row(*fullrow)

            if subject_lower == 'eu documents':
                t_sections['eu_docs'].add_row(*fullrow)

            if subject_lower in ('government motion', 'government motions', 'business motion'):
                t_sections['gov_motions'].add_row(*fullrow)

            if subject_lower in ('general motion'):
                t_sections['gov_motions_gen'].add_row(*fullrow)

            if subject_lower == 'general debate':
                t_sections['gen_debates'].add_row(*fullrow)

            if subject_lower == 'opposition day':
                t_sections['opposition_days'].add_row(*fullrow)

            if subject_lower == 'opposition motion in government time':
                t_sections['oppo_motions_in_gov_time'].add_row(*fullrow)
            if subject_lower == 'backbench business':
                t_sections['backbench_business'].add_row(*fullrow)
            if subject_lower in ('private member\'s motion',
                                 'private member’s motion',
                                 'private members\' motion'):
                t_sections['pm_motion'].add_row(*fullrow)
            if subject_lower == 'ten minute rule motion':
                t_sections['ten_min_motion'].add_row(*fullrow)
            if 'no. 24 debate' in subject_lower:
                t_sections['emergency_debates'].add_row(*fullrow)
            if 'adjournment' in subject_lower:
                t_sections['adjournment_debates'].add_row(*fullrow)
            if subject_lower == 'estimates day':
                t_sections['estimates'].add_row(*fullrow)
            if subject_lower == 'money resolution':
                t_sections['money'].add_row(*fullrow)
            if subject_lower == 'ways and means':
                t_sections['ways_and_means'].add_row(*fullrow)
            if 'affirmative' in subject_lower:
                t_sections['affirmative_sis'].add_row(*fullrow)
            if subject_lower == 'negative statutory instrument':
                t_sections['negative_sis'].add_row(*fullrow)
            if subject_lower == 'questions':
                t_sections['questions'].add_row(*fullrow)
            if subject_lower == 'topical questions':
                t_sections['topical_questions'].add_row(*fullrow)
            if subject_lower in ('urgent question', 'urgent questions'):
                t_sections['urgent_questions'].add_row(*fullrow)
            if subject_lower == 'statement':
                t_sections['statements'].add_row(*fullrow)
            if subject_lower == 'business statement':
                t_sections['business_statements'].add_row(*fullrow)
            if 'committee statement' in subject_lower:
                t_sections['committee_statements'].add_row(*fullrow)
            if 'no. 24 application' in subject_lower:
                t_sections['app_for_emerg_debate'].add_row(*fullrow)
            if subject_lower in ('point of order', 'points of order'):
                t_sections['points_of_order'].add_row(*fullrow)
            if 'public petition' in subject_lower:
                t_sections['public_petitions'].add_row(*fullrow)
            if subject_lower == 'prayers':
                # prayers are not itemised
                # t_sections['prayers'].add_row(*fullrow)
                t_sections['prayers'].duration += entry.duration
                t_sections['prayers'].after_appointed_time += entry.aat

            miscellaneous_options = ('tributes', 'election of a speaker',
                                     'suspension', 'observation of a minute\'s silence',
                                     'personal statement',
                                     'presentation of private members\' bills')
            if (subject_lower in miscellaneous_options
                    or 'message to attend the lords' in subject_lower):

                # for Miscellaneous we will also include stuff in col_subject3
                misc_cells = [
                    forematted_date,
                    ': '.join([entry.subject1, entry.subject2]).rstrip(': '),
                    # formatted_duration,
                    # formatted_a_a_t
                    entry.duration,
                    entry.aat
                ]
                t_sections['miscellaneous'].add_row(misc_cells, entry.duration, entry.aat)

        for table_section in t_sections.values():
            if len(table_section) > 0:
                table_section.add_to(table_ele)

        # need to also add prayers even though there are no rows
        t_sections['prayers'].add_to(table_ele)

        # now create XML for InDesign
        # create root element
        output_root = Element('root')
        output_root.append(table_ele)
        output_tree = etree.ElementTree(output_root)
        output_tree.write(os.path.join(output_folder_path, 'House_Analysis_4.xml'),
                          encoding='UTF-8', xml_declaration=True)

        # also output the contents file
        # CH_AnalysisTableSection.output_contents('CH_contents.txt')

        # build up CH_contents.txt again
        # part_1 duration
        text = f'\tPart 2\t{format_timedelta(CH_AnalysisTableSection.part_dur)}' \
            f'\t{format_timedelta(CH_AnalysisTableSection.part_aat)}'

        previous_number = 0
        for table_section in t_sections.values():
            if table_section.table_num > previous_number:
                previous_number = table_section.table_num
                # if table_section.duration > timedelta(seconds=0)
                # or table_section.after_appointed_time > timedelta(seconds=0):
                table_num = table_section.table_num
                table_num_dur_formatted = format_timedelta(
                    CH_AnalysisTableSection.table_num_dur.get(table_num, timedelta()))
                table_num_aat_formatted = format_timedelta(
                    CH_AnalysisTableSection.table_num_aat.get(table_num, timedelta()))
                text += (f'\n\t{table_num}'
                         f'\t{table_num_dur_formatted}'
                         f'\t{table_num_aat_formatted}')
            else:
                try:
                    title_number = table_section.title.split(":\t")[0]
                except Exception:
                    title_number = table_section.title
                formatted_dur = format_timedelta(table_section.duration)
                formatted_aat = format_timedelta(table_section.after_appointed_time)
                text += f'\n\t{title_number}\t{formatted_dur}\t{formatted_aat}'
        print(text)

    def wh_diary(self, output_folder_path: str = ''):

        self.check_wh()

        table_ele = id_table(
            [('Time', 35), ('Subject', 400), ('Duration', 45)],
            table_class=WH_Diary_Table
        )

        if len(DATE_NUM_LOOK_UP) == 0:
            print('Data for the chamber has not yet been processed so the chamber number will'
                  ' not be put in the westminstar hall table. The square brackets will'
                  ' instead be left blank.')

        wh_data = cast(Worksheet, self.input_workbook[WH_SHEET_TITLE])

        # output = []
        session_total_time = timedelta(seconds=0)
        # day_total_time     = timedelta(seconds=0)

        table_sections = []

        previous_day = 1

        for c, excel_row in enumerate(wh_data.iter_rows()):

            if c == 0:
                continue

            entry = WHRow(excel_row)

            if c == 1:
                chamber_daynum = DATE_NUM_LOOK_UP.get(entry.date, '')
                sec_title = (f'{entry.day}.\u2002[{chamber_daynum}]'
                             f'\u2002{entry.date.strftime("%A %d %B %Y")}')
                table_sections.append(WH_DiaryDay_TableSection(sec_title))


            if entry.day != previous_day:
                previous_day = entry.day

                # day_total_time = timedelta()

                table_sections[-1].add_to(table_ele, session_total_time)

                # chamber_day_number = f'{row_values[0]}.'.strip()

                # if avaliable we will get the day number of the chamber for this date
                chamber_daynum = DATE_NUM_LOOK_UP.get(entry.date, '')
                # actually we will calculate the day number
                # rather than getting the number from excel
                # day_number_counter += 1

                tbl_sec_title_without_date = (f'{entry.day}.\u2002[{chamber_daynum}]'
                                              f'\u2002{entry.date.strftime("%A %d %B %Y")}')
                table_sections.append(WH_DiaryDay_TableSection(tbl_sec_title_without_date))


            # need to add up all the durations
            session_total_time += entry.duration

            # day_total_time += entry.duration

            # there will be 3 cells per row
            cell = ID_Cell()
            # text_for_middle_cell = SubElement(cell, 'text')
            if entry.subject1:
                bold = SubElement(cell, 'Bold')
                bold.text = entry.subject1
                if entry.subject2:
                    bold.tail = f': {entry.subject2}'
                cells = make_id_cells([entry.time.strftime('%H.%M'), cell, entry.duration])  # type: ignore
                table_sections[-1].add_row(cells, entry.duration)

        # last table section will not have been added in the above loop
        table_sections[-1].add_to(table_ele, session_total_time)


        # now create XML for InDesign
        output_root = Element('root')  # create root element
        output_root.append(table_ele)
        tree = etree.ElementTree(output_root)
        tree.write(os.path.join(output_folder_path, 'WH_diary.xml'),
                   encoding='UTF-8', xml_declaration=True)

    def wh_analysis(self, output_folder_path: str = ''):

        wh_data = cast(Worksheet, self.input_workbook['Westminster Hall'])

        # add a new table element with headings
        table_ele = id_table([('Date', 95), ('Detail', 340), ('Duration', 45)],
                             table_class=WH_Table)

        # can now use dict (rather than ordered dict) as order is guaranteed
        t_sections = {
            # the order matters!
            'private': WH_AnalysisTableSection(
                '1a:\tPrivate Members’ Debates',
                'WH1 Members debates',
                1),
            'bbcom': WH_AnalysisTableSection(
                '1b:\tPrivate Members’ (Backbench Business Committee recommended) Debates',
                'WH2 BBCom debates',
                1),
            'liaison': WH_AnalysisTableSection(
                '2:\tLiaison Committee Debates',
                'WH3 Liaison Com debates',
                2),
            'e_petition': WH_AnalysisTableSection(
                '3:\tDebates on e-Petitions',
                'WH4 e-Petitions',
                3),
            'suspension': WH_AnalysisTableSection(
                '4:\tSuspensions',
                'WH5 Suspensions',
                4),
            'miscellaneous': WH_AnalysisTableSection(
                '5:\tMiscellaneous',
                'WH6 Miscellaneous',
                5),
            'statements': WH_AnalysisTableSection(
                '6:\tStatements',
                'WH7 Statements',
                6),
            'brexit': WH_AnalysisTableSection(
                '7:\tTime spent on Brexit',
                'Brexit',
                7)
        }

        for c, excel_row in enumerate(wh_data.iter_rows()):
            if c == 0:
                continue
            # if c > 1500:
            #     # for the moment lets just work with a quarter of the content
            #     break

            # excel_row = [item.value for item in row[:8]]  # only interested in the first few cells

            # check to see if all items in list are None as there are lots of blank rows
            if all(bool(v) is False for v in excel_row[:8]):
                continue

            entry = WHRow(excel_row)

            # if not isinstance(row_values[2], time):
            #     # assume we can skip any records without a time
            #     continue
            # if row_values[1] == 'Date' and row_values[2] == 'Time':
            #     # ignore the row with the date in
            #     continue

            # if row_values[4] in ('Daily totals', 'Totals for Session'):
            #     continue
            # if not row_values[0] or isinstance(row_values[0], int):
            #     # some values are blank or are numbers
            #     continue

            # sometimes the value in the cell is not a time but is instead a datetime e.g. cell H751
            # for j in (2, 7):
            #     if isinstance(row_values[j], datetime):
            #         time_obj = row_values[j].time()
            #         print(f'There is a datetime at row {c + 1}:', str(row_values[j]))
            #         print(f'This has been converted to the following time: {time_obj}')
            #         row_values[j] = time_obj

            # for i, value in enumerate(row_values):
            #     if value is None:
            #         row_values[i] = ''
            #     if isinstance(value, time):
            #         row_values[i] = value.strftime('%H.%M')


            # if not isinstance(row_values[1], datetime):
            #     print(f'Error in row {c}, expected datetime but got, {row_values[0]}')
            # elif not isinstance(row_values[3], str):
            #     print(f'Error in row {c}, expected str but got, {row_values[3]}')
            # else:
            #     forematted_date = format_date(row_values[1])

            forematted_date = format_date(entry.date)  # type: ignore

            # cells = [forematted_date, row_values[4], row_values[7]]
            # col_3_val = row_values[3].strip()
            # col_4_val = row_values[4].strip()
            # hours_mins = row_values[7].split('.')
            # if len(hours_mins) == 2:
            #     duration = timedelta(hours=int(hours_mins[0]), minutes=int(hours_mins[1]))
            # else:
            #     duration = timedelta()
            cells_vals = [
                forematted_date,
                entry.subject2,
                entry.duration,
            ]
            fullrow = [cells_vals, entry.duration]
            if entry.subject1 in ('Debate (Private Member’s)', 'Debate (Private Member\'s)'):
                t_sections['private'].add_row(*fullrow)
            elif entry.subject1 in ('Debate (BBCom recommended)',
                                    'Debate (BBCom)', 'Debate (BBBCom)'):
                t_sections['bbcom'].add_row(*fullrow)
                # bbcom_duration += duration
                # bbcom_cells.extend(cells)
            elif entry.subject1 in ('Debate (Liaison Committee)', ):
                t_sections['liaison'].add_row(*fullrow)
                # liaison_duration += duration
                # liaison_cells.extend(cells)
            elif entry.subject1 in ('Petition', 'Petitions'):
                t_sections['e_petition'].add_row(*fullrow)
                # e_petition_duration += duration
                # e_petition_cells.extend(cells)
            elif entry.subject1 in ('Suspension',) and entry.tags not in ('[Questions]', '[Question]'):
                t_sections['suspension'].add_row(*fullrow)
                # suspension_duration += duration
                # suspension_cells.extend(cells)
            elif entry.subject1 in ('Committee Statement',):
                t_sections['statements'].add_row(*fullrow)
            elif entry.subject1 in ('Time limit', 'Time Limit',
                                    'Observation of a period of silence'):
                t_sections['miscellaneous'].add_row(*fullrow)
                # miscellaneous_duration += duration
                # miscellaneous.extend(cells)
            # if entry.subject1 == '[exit]':
            #     t_sections['brexit'].add_row(cells, duration)


        for table_section in t_sections.values():
            if len(table_section) > 0:
                table_section.add_to(table_ele)

        # create root element
        output_root = Element('root')
        output_root.append(table_ele)
        tree = etree.ElementTree(output_root)
        tree.write(os.path.join(output_folder_path, 'WH_Analysis.xml'),
                   encoding='UTF-8', xml_declaration=True)

        # WH_AnalysisTableSection.output_contents('WH_contents.txt')

        # also output the contents file
        # CH_AnalysisTableSection.output_contents('CH_contents.txt')

        # build up CH_contents.txt again
        # part_1 duration
        text = f'\tPart 4\t{format_timedelta(WH_AnalysisTableSection.part_dur)}'

        previous_number = 0
        for table_section in t_sections.values():
            if table_section.table_num > previous_number:
                previous_number = table_section.table_num

                table_num = table_section.table_num
                table_num_dur_formatted = format_timedelta(
                    WH_AnalysisTableSection.table_num_dur.get(table_num, timedelta()))
                text += (f'\n\t{table_num}'
                         f'\t{table_num_dur_formatted}')
            else:
                try:
                    title_number = table_section.title.split(":\t")[0]
                except Exception:
                    title_number = table_section.title
                formatted_dur = format_timedelta(table_section.duration)
                text += f'\n\t{title_number}\t{formatted_dur}'
        print(text)


def main():

    if len(sys.argv) > 1:
        # do cmd line version
        parser = argparse.ArgumentParser(
            description='Process Sessional diary Excel and create XML for InDesign')

        parser.add_argument('input', metavar='input_file', type=open,
                            help='File path to the Excel file you wish to process. '
                                 'If there are spaces in the path you must use quotes.')

        parser.add_argument('--no-excel',
                            action='store_true',
                            help='Use this flag if you want do not want to output an excel file.')

        parser.add_argument('--include-only',
                            type=str,
                            choices=['chamber', 'wh'],
                            help='Use this option if you want to include *only* '
                                 'one section (e.g. just the Chamber section) '
                                 'rather than both sections')

        args = parser.parse_args(sys.argv[1:])

        if args.include_only == 'chamber':
            run(args.input.name, include_wh=False, no_excel=args.no_excel)
        elif args.include_only == 'wh':
            run(args.input.name, include_chamber=False, no_excel=args.no_excel)
        else:
            run(args.input.name, no_excel=args.no_excel)

    else:
        # run the GUI version
        from package import gui
        gui.mainloop(run_callback=run)


def run(excel_file_path: str,
        output_folder_path: str = '',
        include_chamber=True,
        include_wh=True,
        no_excel=False):

    if not output_folder_path:
        output_folder_path = os.path.dirname(excel_file_path)

    # sessional_diary = Sessional_Diary('Sessional diary 2019-21_downloaded_2021-06-18.xlsx')
    sessional_diary = Sessional_Diary(excel_file_path, no_excel)

    if include_chamber:
        # create house diary
        sessional_diary.house_diary(output_folder_path)

        # create house analysis
        sessional_diary.house_analysis(output_folder_path)

    if include_wh:
        # crete Westminster hall diary
        sessional_diary.wh_diary(output_folder_path)

        # create Westminster hall analysis
        sessional_diary.wh_analysis(output_folder_path)

    # remove the default sheet
    if Excel.out_wb is not None:
        del Excel.out_wb['Sheet']
        Excel.out_wb.save(filename=os.path.join(output_folder_path, 'Analysis.xlsx'))


def id_table(list_of_tuples: list[tuple[str, int]],
             table_class: Type[WH_Table]):
    """Takes a list of 2 tuples of table header and cell widths"""

    # table element
    # the table_class is based on etree.ElementBase
    # name of the xml element will default to the name of the class
    # class can be instantiated despite what the type checker may think
    table_ele = table_class(  # type: ignore
        nsmap=NS_MAP,
        attrib={AID + 'table': 'table',
                AID + 'tcols': str(len(list_of_tuples)),
                AID5 + 'tablestyle': 'Part1Table'})
    # table_ele.tag = 'Table'
    table_ele.increment_rows()

    # add heading elements to table
    for item in list_of_tuples:
        heading = SubElement(table_ele, 'Cell',
                             attrib={AID + 'table': 'cell',
                                     AID + 'theader': '',
                                     AID + 'ccolwidth': str(item[1])})
        heading.text = item[0]

    return table_ele


if __name__ == '__main__':
    main()
