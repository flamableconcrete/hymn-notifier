import calendar
import os
from pprint import pprint
from textwrap import dedent

import pygsheets

from delorean import Delorean, parse
from dotenv import load_dotenv


def main():
    hymn_spreadsheet = os.getenv('HYMN_SPREADSHEET')
    hymn_worksheet = os.getenv('HYMN_WORKSHEET')
    credentials_file = os.getenv("GOOGLE_OAUTH_CREDENTIALS_FILE")
    signature = os.getenv("SIGNATURE")

    gc = pygsheets.authorize(client_secret=credentials_file)

    # Open the hymns spreadsheet
    wkbk = gc.open(hymn_spreadsheet)
    sacrament_hymn_sheet = wkbk.worksheet_by_title(hymn_worksheet)

    next_sunday = Delorean().next_sunday().truncate('day')

    all_records = sacrament_hymn_sheet.get_all_records()
    msg = None
    for record in all_records:
        date = record['Date']
        row_date = parse(date, dayfirst=False, yearfirst=False)
        if next_sunday == row_date:
            # print(row_date, '---', record['Opening Hymn'])
            month = calendar.month_name[next_sunday.date.month]
            day = next_sunday.date.day

            opening_hymn = record['Opening Hymn']
            sacrament_hymn = record['Sacrament Hymn']
            intermediate_hymn = record['Intermediate Hymn']
            closing_hymn = record['Closing Hymn']

            msg = dedent(f"""\
            Here are the hymns for {month} {day}.

            Opening: {opening_hymn}
            Sacrament: {sacrament_hymn}
            Intermediate: {intermediate_hymn}
            Closing: {closing_hymn}

            Let me know if you have any questions or would like to make any changes/suggestions. Thanks!

            {signature}
            """)

            if opening_hymn.lower() == 'n/a':
                msg = f'There are no hymns for {month} {day}'
    print(msg)


if __name__ == "__main__":
    load_dotenv()
    main()
