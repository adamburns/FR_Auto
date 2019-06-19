import xlrd
import pandas as pd

#TODO(Adam): Add a file to read in
#monthly costs below:
#monthly = ('Source Doc. - Report Costs for FactRight Tab of March Financial Review.xlsx')
#monthly revenue below:
#monthly = ('Source Doc. - Report Revenue for FactRight Tab of March Financial Review.xlsx')
#monthly consolidated below:
monthly = ('Source Doc. - Final Consolidated Financials March 2019.xlsx')

wb = xlrd.open_workbook(monthly)
#open the first sheet in the workbook
sheet = wb.sheet_by_index(0)
print([sheet.cell(0, col_index).value for col_index in range(sheet.ncols) if sheet.cell(0, col_index).value])


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


def check_data_node(sheet, loc, cols):
    data = [sheet.cell(loc[0], col).value for col in cols if sheet.cell(loc[0], col).value]
    if not data:
        return False
    return True


def check_next_node(sheet, loc, cols):
    if check_closing_node(sheet, loc):
        return 2
    elif check_neutral_node(sheet, loc):
        return 3
    elif check_child_node(sheet, loc):
        return 1
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


def max_rows(sheet):
    return sheet.nrows


def max_cols(sheet):
    return sheet.ncols


def get_cols(sheet):
    cols = [col_index for col_index in range(sheet.ncols) if sheet.cell(0, col_index).value]
    print('test')
    print(cols)
    return cols


def get_headers(sheet, cols):
    headers = []
    for loc in cols:
        headers.append(sheet.cell_value(0, loc))
    return headers


def check_account(sheet, loc):
    print(sheet.cell_value(loc[0], loc[1]))
    current_account = sheet.cell_value(loc[0], loc[1]).split(" ")[0]
    if current_account.isdigit():
        return True
    return False


def get_account(sheet, loc):
    current_account = sheet.cell_value(loc[0], loc[1]).split(" ")[0]
    return current_account


def check_account_total(sheet, loc):
    current_account = sheet.cell_value(loc[0], loc[1]).split(" ")[0]
    if current_account is 'Total':
        return True
    return False


def scan_across_doc(sheet, loc, cols, headers, account):
    col_vals = [account] + [sheet.cell_value(loc[0], col) for col in cols]
    print(headers)
    print(col_vals)
    col_vals = [col_vals]
    df = pd.DataFrame(col_vals, columns=headers)
    print(df)
    return df


def scan_down_doc(sheet):
    print(max_rows(sheet))
    cursor = [0, 0]
    cols = get_cols(sheet)
    headers = ['Account'] + get_headers(sheet, cols)
    df = pd.DataFrame(columns=headers)
    account = 0
    while(cursor[0] < max_rows(sheet) - 1):      
        print(str(cursor[0]) + '/' + str(max_rows(sheet)))  
        next_node = check_next_node(sheet, cursor, cols)
        print(next_node)
        if next_node == 1:
            #Get parent node and add to tree
            if not check_data_node(sheet, cursor, cols):
                cursor = get_next_node(next_node, cursor)
                if check_account(sheet, cursor):
                    account = get_account(sheet, cursor)
                continue
            else:
                cursor = get_next_node(next_node, cursor)
                if check_account(sheet, cursor):
                    account = get_account(sheet, cursor)
                    df = df.append(scan_across_doc(sheet, cursor, cols, headers, account))
                else:
                    continue
        elif next_node == 2:
            if check_account_total(sheet, cursor):
                account = 0
            cursor = get_next_node(next_node, cursor)
        elif next_node == 3:
            if not check_data_node(sheet, cursor, cols):
                cursor = get_next_node(next_node, cursor)
                if check_account(sheet, cursor):
                    account = get_account(sheet, cursor)
                continue
            cursor = get_next_node(next_node, cursor)
            df = df.append(scan_across_doc(sheet, cursor, cols, headers, account))
        elif next_node == 4:
            cursor = get_next_node(next_node, cursor)
            df = df.append(scan_across_doc(sheet, cursor, cols, headers, account))
        else:
            return df
    return df


print_df = scan_down_doc(sheet)
print(print_df.reset_index())
