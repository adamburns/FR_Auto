from anytree import Node, RenderTree
import xlrd

monthly = ('Source Doc. - Final Consolidated Financials March 2019.xlsx')
wb = xlrd.open_workbook(monthly)
sheet = wb.sheet_by_index(1)


def max_columns(sheet):
    return sheet.ncols


def max_rows(sheet):
    return sheet.nrows


def get_cols(sheet):
    cols = [col_index for col_index in range(sheet.ncols) 
            if sheet.cell(0, col_index).value]
    return cols


def within_bounds(loc, max_columns, max_rows):
    if((loc[0] < max_rows) and (loc[1] < max_columns)):
        return True
    else:
        return False


def is_account(sheet, loc):
    print('is_account: \t' + str(loc))
    account = sheet.cell_value(*loc)
    if account.isdigit():
        return True
    return False


def get_account(sheet, loc):
    account =  sheet.cell_value(*loc).split(" ")[0]
    return account


def is_total_account(sheet, loc):
    total = sheet.cell_value(*loc).split(" ")[0]
    account = sheet.cell_value(*loc).split(" ")[1]
    if (total is 'Total') and (account.isdigit()):
        return True
    return False


def pendulum_search(sheet, loc, first_column):
    #check column
    print('pendulum search: \t' + str(loc))
    print('first column: \t' + str(first_column))
    loc = loc[0]
    print('pendulum search mash: \t' + str(loc))
    loc_right = [loc[0] + 1, loc[1] + 1]
    loc_down = [loc[0] + 1, loc[1]]
    loc_left = [loc[0] + 1, loc[1] - 1]
    
    if(loc[1] == first_column):
        return pendulum_search(sheet, [loc[0], loc[1] - 1])
    if(loc[1] + 1 != first_column):
        if sheet.cell_value(*loc_right):
            return [loc_right, 0]
    if sheet.cell_value(*loc_down):
        return [loc_down, 1]
    if sheet.cell_value(*loc_left):
        return [loc_left, 2]
    return [[max_rows(sheet), max_columns(sheet)], 0]


def scan_down_doc(sheet):
    RIGHT = 0
    DOWN = 1
    LEFT = 2
    loc = [[1, 2], 0]
    next_loc = [[1, 2], 0]
    cols = get_cols(sheet)
    max_cs = max_columns(sheet)
    max_rs = max_rows(sheet)
    accounts = dict()
    parent_account = None
    while(within_bounds(loc[0], max_cs, max_rs)):
        #when step left, adjust current_account
        loc = next_loc
        next_loc = pendulum_search(sheet, loc, cols[0])
        if(not within_bounds(next_loc[0], max_cs, max_rs)):
            break
        if next_loc[1] == RIGHT:
            print('if right:\t' + str(next_loc))
            if(is_account(sheet, loc[0]) and is_account(sheet, next_loc[0])):
                parent_account = get_account(sheet, *loc[0])
                current_account = get_account(sheet, *next_loc[0])
                Node(get_account(sheet, next_loc[0]), parent = accounts[parent_account])
            elif(is_account(sheet, next_loc[0])):
                Node(get_account(sheet, next_loc[0]), parent = accounts[parent_account])
            elif(is_account(sheet, loc[0])):
                parent_account = get_account(sheet, loc[0])
            else:
                continue
        if next_loc[1] == DOWN:
            if(is_account(sheet, loc[0]) and is_account(sheet, next_loc[0])):
                Node(get_account(sheet, next_loc[0]), parent = accounts[parent_account])
            elif(is_account(sheet, next_loc[0])):
                Node(get_account(sheet, next_loc[0]), parent = accounts[parent_account])
            else:
                continue
        if next_loc[1] == LEFT:
            if(is_total_account(sheet, next_loc[0])):
                parent_account = parent_account.parent
            elif(is_account(sheet, loc[0])):
                parent_account = accounts[get_account(sheet, loc[0])].parent
            else:
                continue
    return accounts

test = scan_down_doc(sheet)
for k, v in test.items():
    print(k, v)
