from bs4 import BeautifulSoup
import os
import openpyxl
import requests


def filter_links(URL):
    return


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    wbkName = 'org_chart.xlsx'
    wrkbk = openpyxl.load_workbook(wbkName)
    sheet = wrkbk.get_sheet_by_name('org_chart')
    for i in range(2, sheet.max_row + 1):
        cell_obj = sheet.cell(row=i, column=1)
        url = str(cell_obj.value)
        req = requests.get(url, 'html.parser')
        if(req.text[40:45] == 'Error'):
            coords = "C" + str(i)
            sheet[coords] = "Access Denied"

    wrkbk.save(wbkName)
    wrkbk.close

