import calendar
import os
from pprint import pprint
from textwrap import dedent

import gspread

from delorean import Delorean, parse
from dotenv import load_dotenv
from oauth2client.service_account import ServiceAccountCredentials

HYMN_SPREADSHEET = "Centreville First Ward Sacrament Hymns"


def parse_num(hymn_cell):
    return hymn_cell.split()[0]


def parse_hymn(hymn_cell):
    return hymn_cell[hymn_cell.find(' ') + 1:]


def main():
    scope = ['https://www.googleapis.com/auth/drive']
    svc_account_credentials = os.getenv("GOOGLE_SERVICE_ACCOUNT_CREDENTIALS_FILE")

    credentials = ServiceAccountCredentials.from_json_keyfile_name(svc_account_credentials, scope)
    gc = gspread.authorize(credentials)

    # Open a worksheet from spreadsheet with one shot
    wkbk = gc.open(HYMN_SPREADSHEET)

    sacrament_hymn_sheet = wkbk.worksheet("Sacrament Hymns")

    # Fetch a cell range
    # cell_list = sacrament_meeting.range('A1:T200')
    # for cell in cell_list:
    #     print('cell: ', type(cell), cell)

    all_rows = sacrament_hymn_sheet.get_all_values()
    next_sunday = Delorean().next_sunday().truncate('day')
    row_info = None
    # pprint(all_rows)
    for row in all_rows:
        date = row[0]

        # skip header row
        if 'Date' in date:
            continue

        row_date = parse(date)
        if next_sunday != row_date:
            continue

        # print(f'Next Sunday: {next_sunday} ?==? row date: {row_date}')
        print(row)
        row_info = row

    month = calendar.month_name[next_sunday.date.month]
    day = next_sunday.date.day
    # opening_num = parse_num(row_info[2])
    # opening_hymn = parse_hymn(row_info[2])
    # sacrament_num = parse_num(row_info[3])
    # sacrament_hymn = parse_hymn(row_info[3])
    # intermediate_num = parse_num(row_info[4])
    # intermediate_hymn = parse_hymn(row_info[4])
    # closing_num = parse_num(row_info[5])
    # closing_hymn = parse_hymn(row_info[5])

    opening_hymn = row_info[2]
    sacrament_hymn = row_info[3]
    intermediate_hymn = row_info[4]
    closing_hymn = row_info[5]
    signature = os.getenv("SIGNATURE")

    msg = dedent(f"""\
    Here are the hymns for {month} {day}.

    Opening: {opening_hymn}
    Sacrament: {sacrament_hymn}
    Intermediate: {intermediate_hymn}
    Closing: {closing_hymn}

    Let me know if you have any questions or would like to make any changes/suggestions. Thanks!

    {signature}
    """)
    print(msg)


if __name__ == "__main__":
    load_dotenv()
    main()
