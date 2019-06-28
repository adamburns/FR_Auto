from anytree import Node, RenderTree
import xlrd


monthly = ('Source Doc. - Final Consolidated Financials March 2019.xlsx')
wb = xlrd.open_workbook(monthly)
sheet = wb.sheet_by_index(1)


def check_child_node(sheet, loc):
    loc = [axis + 1 for axis in loc]
    if not sheet.cell_value(*loc):
        return False
    return True


def check_account_node(sheet, loc):
    try:
        int(sheet.cell_value(*loc).split(" ")[0])
        return True
    except ValueError:
        return False


def check_child_account(sheet, loc):
    current_node = loc
    child_node = [axis + 1 for axis in loc]
    if(check_child_node(*loc)):
        child_node = [axis + 1 for axis in loc]
        if(check_account_node(*child_node)):
            return child_node
        #CHECK IF CHILD_NODE COLX IS IN BOUNDS
        elif(IN COLUMN BOUNDS):
            return check_child_account(sheet, child_node)
    return None


def last_child(sheet, loc):
    for y in range(loc[0], max_rows(sheet)):
        

def get_account(sheet, loc):
    account = int(sheet.cell_value(*loc).split(" ")[0])
    return account


def find_all_child_nodes():


def search_next_node(sheet, loc, cols, tree):
    if(loc[0] >= (max_rows(sheet) - 2)):
        return #STOP
    for y in range(1, 2):
        for x in range(cols[0]):
            cursor = (loc[0] + y, x)
            if(check_account_node(cursor)):
                #check child nodes
                if(check_child_account(sheet, cursor)):
                    parent_account = str(get_account(sheet, cursor))
                    tree[parent_account] = Node(parent_account)
                    for z in range(last_child(sheet, loc)):
                        child_cursor = [cursor[0] + z + 1, cursor[1] + 1]
                        child_account = str(get_account(sheet, child_cursor))
                        Node(child_account, parent = tree[parent_account])
                        
    #return get_next_node(check_next_node(sheet, loc, cols), loc)
    if(loc[0] < max_rows(sheet)):
        return search_next_node(sheet, [loc[0] + 2, loc[1]], cols)
    return [max_rows(sheet), loc[1]]

dictionary = dict()
parent = "5100"
children = ["5113", "5112", "5117"]
dictionary["5100"] = Node("5100")
for child in children:
    Node(child, parent = dictionary[parent])

print(RenderTree(dictionary["5100"]))
