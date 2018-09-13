import os
from pprint import pprint
from textwrap import dedent

import gspread

from delorean import Delorean
from dotenv import load_dotenv
from oauth2client.service_account import ServiceAccountCredentials

HYMN_SPREADSHEET = "Centreville First Ward Sacrament Hymns"


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

    cell_list = sacrament_hymn_sheet.get_all_values()
    d = Delorean()
    d.next_sunday()
    d.truncate('day')
    # pprint(cell_list)
    print(cell_list[28])

    # date = sacrament_meeting.col_values(2)
    # pprint(date)

    month = None
    day = None
    opening_num = None
    opening_hymn = None

    sacrament_num = None
    sacrament_hymn = None

    intermediate_num = None
    intermediate_hymn = None

    closing_num = None
    closing_hymn = None

    msg = dedent(f"""Here are the hymns for {month} {day}.

                     Opening: {opening_num} {opening_hymn}
                     Sacrament: {sacrament_num} {sacrament_hymn}
                     Intermediate: {intermediate_num} {intermediate_hymn}
                     Closing: {closing_num} {closing_hymn}

                     Let me know if you have any questions or would like to make any changes/suggestions. Thanks!

                     Brother Ondricek
                     """)


if __name__ == "__main__":
    load_dotenv()
    main()
