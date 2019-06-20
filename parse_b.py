import xlrd
import pandas as pd

#TODO(Adam): Add a file to read in
#monthly costs below:
#monthly = ('Source Doc. - Report Costs for FactRight Tab of March Financial Review.xlsx')
#monthly revenue below:
#monthly = ('Source Doc. - Report Revenue for FactRight Tab of March Financial Review.xlsx')
#monthly consolidated below:
#monthly = ('Source Doc. - Final Consolidated Financials March 2019.xlsx')
#budget details below:
budget = ('Source Doc. - Budget Template -2019.xlsx')
wb = xlrd.open_workbook(budget)
sheet = wb.sheet_by_index(10)
SEARCH_RADIUS = 3

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


def check_reach_node(sheet, loc):
    if not sheet.cell_value(loc[0] + 2, loc[1]):
        return False
    return True


def check_data_node(sheet, loc, cols):
    data = [sheet.cell(loc[0], col).value for col in cols if sheet.cell(loc[0], col).value]
    print(data)
    if not data:
        return False
    return True


def check_next_node(sheet, loc, cols):
    if check_neutral_node(sheet, loc):
        return 3
    elif check_closing_node(sheet, loc):
        return 2
    elif check_child_node(sheet, loc):
        return 1
    elif check_reach_node(sheet, loc):
        return 6
    elif check_data_node(sheet, loc, cols):
        return 4
    else:
        return 5


def get_next_node(arg, loc):
    if arg == 1:
        return [loc[0] + 1, loc[1] + 1]
    if arg == 2:
        return [loc[0] + 1, loc[1] - 1]
    if arg in {3, 4}:
        return [loc[0] + 1, loc[1]]
    if arg == 6:
        return [loc[0] + 2, loc[1]]


def search_next_node(sheet, loc, cols):
    global SEARCH_RADIUS
    if(loc[0] >= (max_rows(sheet) - 2)):
        return [max_rows(sheet) - 1, loc[1]]
    for y in range(1, SEARCH_RADIUS):
        for x in range(cols[0] - 1):
            cursor = (loc[0] + y, x)
            try:
                int(sheet.cell(*cursor).value)
                return cursor
            except ValueError:
                continue
    #return get_next_node(check_next_node(sheet, loc, cols), loc)
    if(loc[0] < max_rows(sheet)):
        return search_next_node(sheet, [loc[0] + 2, loc[1]], cols)
    return [max_rows(sheet), loc[1]]


def max_rows(sheet):
    return sheet.nrows


def max_cols(sheet):
    return sheet.ncols


def get_cols(sheet):
    cols = [col_index for col_index in range(sheet.ncols) if sheet.cell(4, col_index).value]
    return cols

def get_headers(sheet, cols):
    headers = []
    for loc in cols:
        headers.append(sheet.cell_value(0, loc))
    return headers


def check_account(sheet, loc):
    current_account = sheet.cell_value(loc[0], loc[1])
    try:
        current_account = int(current_account)
        return True
    except ValueError:
        return False


def get_account(sheet, loc):
    current_account = round(sheet.cell_value(*loc))
    return current_account


def check_account_total(sheet, loc):
    current_account = sheet.cell_value(loc[0], loc[1]).split(" ")[0]
    if current_account is 'Total':
        return True
    return False


def scan_across_doc(sheet, loc, cols, headers, account):
    col_vals = [account] + [sheet.cell_value(loc[0], col) for col in cols]
    col_vals = [col_vals]
    df = pd.DataFrame(col_vals, columns=headers)
    return df


def scan_down_doc(sheet):
    print(max_rows(sheet))
    cursor = [4, 0]
    cols = get_cols(sheet)
    headers = ['Account'] + get_headers(sheet, cols)
    df = pd.DataFrame(columns=headers)
    account = 0
    while True:      
        print(str(cursor[0]))  
        cursor = search_next_node(sheet, cursor, cols)
        if(cursor[0] >= (max_rows(sheet) - 2)):
            break
        account = get_account(sheet, cursor)
        df = df.append(scan_across_doc(sheet, cursor, cols, headers, account))
    return df

#print(get_cols(sheet))
print_df = scan_down_doc(sheet)
print_df.reset_index()
print(print_df)
#test = (27, 0)
#print(sheet.cell(*test).value)
#print(search_next_node(sheet, test, get_cols(sheet)))
