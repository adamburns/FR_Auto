import xlrd
import pandas as pd

#TODO(Adam): Add a file to read in
monthly = ("Source Doc. - Report Costs for FactRight Tab of March Financial Review.xlsx")

wb = xlrd.open_workbook(monthly)
#open the first sheet in the workbook
sheet = wb.sheet_by_index(1)
print(sheet)

def check_child_account(sheet, loc)
    #If no child account is found, return current location
    if not sheet.cell_value(loc[0] + 1, loc[1] + 1):
        return False
    #Increment each coordinate by one if Child is found
    return True
    #return [x + 1 for x in loc]

def max_rows(sheet):
    return sheet.nrows


def max_cols(sheet):
    return sheet.ncols

def get_cols(sheet, loc):
    cols = []
    while(loc[1] < max_cols(sheet)):
        if not sheet.cell_value(0, loc[1]):
            loc[1] += 1
            continue
        cols.append(loc[1])
    return cols


def get_headers(sheet, cols):
    headers = []
    for loc in cols:
        headers.append(sheet.cell_value(0, loc))
    return headers


def scan_down_doc(sheet)
    cursor = [0, 0]
    cols = get_cols(sheet, cursor)
    headers = ['Account'] + get_headers()
    df = pd.DataFrame(columns=headers)
    
    while(cursor[0] < max_rows(sheet)):
        current_account = sheet.cell_value(cursor[0], cursor[1]).split(" ")[0]
        if current_account.isdigit():
            #create account
            
            #gather data

def scan_across_doc(sheet, loc, df)
    
