import xlrd
import pandas as pd

#TODO(Adam): Add a file to read in
monthly = ("Source Doc. - Report Costs for FactRight Tab of March Financial Review.xlsx")

wb = xlrd.open_workbook(monthly)
#open the first sheet in the workbook
sheet = wb.sheet_by_index(1)
print(sheet)


def check_child_node(sheet, loc):
    #If no child account is found, return current location
    if not sheet.cell_value(loc[0] + 1, loc[1] + 1):
        return False
    #Increment each coordinate by one if Child is found
    return True
    #return [x + 1 for x in loc]


def get_parent_node(sheet, loc):
    return [loc[0] - 1, loc[1] - 1]


def check_closing_node(sheet, loc):
    if not sheet.cell_value(loc[0] + 1, loc[1] - 1):
        return False
    return True


def check_neutral_node(sheet, loc):
    if not sheet.cell_value(loc[0] + 1, loc[1]):
        return False
    return True


def check_next_node(sheet, loc):
    if check_child_node(sheet, loc):
        return 1
    elif check_closing_node(sheet, loc):
        return 2
    elif check_neutral_node(sheet, loc):
        return 3
    else:
        return 4


def get_next_node(arg, loc):
    if arg == 1:
        return [loc[0] + 1, loc[1] + 1]
    if arg == 2:
        return [loc[0] + 1, loc[1] - 1]
    if arg == 3:
        return [loc[0] + 1, loc[1]]


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


def check_account(sheet, loc):
    current_account = sheet.cell_value(loc[0], loc[1]).split(" ")[0]
    if current_account.isdigit():
        return True
    return False


def get_account(sheet, loc):
    current_account = sheet.cell_value(loc[0], loc[1]).split(" ")[0]
    return current_account


def scan_across_doc(sheet, loc, cols, df, account):
    


def scan_down_doc(sheet):
    cursor = [0, 0]
    cols = get_cols(sheet, cursor)
    headers = ['Account'] + get_headers()
    df = pd.DataFrame(columns=headers)
    
    while(cursor[0] < max_rows(sheet)):        
        next_node = check_next_node(sheet, cursor)
        if next_node == 1:
            #Get parent node and add to tree
            cursor = get_next_node(next_node, cursor)
            if check_account(sheet, cursor):
                account = get_account(sheet, cursor)
                scan_across_doc(sheet, cursor, cols, df, account)
        if next_node == 2 or next_node == 3:
            cursor = get_next_node(next_node, cursor)
        if next_node == 4:
            break


def scan_across_doc(sheet, loc, df, account)
    #given get_cols(), use those data points with the current row and add that to the df    
