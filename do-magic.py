import os
import warnings
from tqdm import tqdm
import string
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from openpyxl.styles import Font

warnings.simplefilter("ignore")


def set_border(ws, cell_range):
    thin = openpyxl.styles.Side(border_style="thin", color="757171")
    for row in ws[cell_range]:
        for cell in row:
            cell.border = openpyxl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
            cell.comment = None
            # cell.alignment = Alignment(wrap_text=True)


def set_header(ws, cell_range):
    for row in ws[cell_range]:
        for cell in row:
            cell.fill = openpyxl.styles.PatternFill(start_color="0B5394", fill_type="solid")
            cell.font = cell.font.copy(color="ffffff")
            cell.alignment = openpyxl.styles.Alignment(horizontal='center')


def check_end(ws):
    no_end = True
    i = 8
    while no_end:
        if ws['A{}'.format(i)].value:
            pass
        else:
            no_end = False
            return i - 1
        i += 1


def check_max_col(ws):
    no_end = True
    i = 8
    while no_end:
        if not ws['{}4'.format(get_column_letter(i))].value:
            no_end = False
        i += 1
    return get_column_letter(i - 2)


def check_start_a(ws):
    for i in range(5, 19):
        if ws['a{}'.format(i)].value:
            return (i)


def check_start_f(ws):
    col_range = list(string.ascii_lowercase)
    for i, n in enumerate(col_range):
        if ws['{}2'.format(n)].value:
            return col_range[i], col_range[i - 1], col_range[i + 1]

def check_for_hide_colums(n, e, ws):
    to_hide = []
    col_range = list(string.ascii_lowercase)
    for i, col in enumerate(col_range):
        if '[Attr' in str(ws['{}{}'.format(col,n)].value):
            to_hide.append(col)
    return(to_hide)


def hide_cols(to_hide, ws):
    for col in to_hide:
        ws.column_dimensions[col.upper()].hidden = True       
    return ws


def check_for_hide_rows(n, m, ws):
    to_hide = []
    for row in range(1, n):
        if '[Attr' in str(ws['{}{}'.format(m, row)].value):
            to_hide.append(row)
    return to_hide


def hide_rows(to_hide, ws):
    for row in to_hide: 
        ws.row_dimensions[row].hidden = True     
    return ws    


def format_first_type(ws):
    default_column_width = 25
    for i in list(string.ascii_lowercase):
        ws.column_dimensions[i].width = default_column_width
    n = check_start_a(ws) # start of first block
    m, e , x = check_start_f(ws) # m = start column of second block; e = end column of first block; x = freeze column
    end = check_end(ws) # end of first block
    col = check_max_col(ws) # end column
    to_hide = check_for_hide_colums(n, e, ws)
    ws = hide_cols(to_hide, ws)
    to_hide = check_for_hide_rows(n, m, ws)
    ws = hide_rows(to_hide, ws)    
    set_border(ws, 'A{}:{}{}'.format(n, e, end))
    set_border(ws, '{}1:{}{}'.format(m, col, end))
    set_header(ws, 'A{}:{}{}'.format(n, e, n))
    set_header(ws, '{}1:{}{}'.format(m, m, n - 1))
    ws['A1'].alignment = Alignment(horizontal='center')
    ws['A1'].alignment = Alignment(wrap_text=True)
    ws['A1'].font = Font(size="9", bold=True, italic=True)
    ws.merge_cells('A1:{}{}'.format(e, n -1))
    ws.freeze_panes = '{}{}'.format(x, n+1)

    return ws


def next_alpha(s):
    return chr((ord(s.upper())+1 - 65) % 26 + 65)



def format_information_result(ws, last_tab):
    default_column_width = 20
    
    for i in list(string.ascii_lowercase):
        ws.column_dimensions[i].width = default_column_width
        if ws['a1'].value == "ID" and i == "a":
            ws.column_dimensions[i].width = "3"
    ws.column_dimensions[last_tab].width = 3
    ws.merge_cells('{}1:{}1'.format(last_tab, next_alpha(last_tab)))
    set_header(ws, 'A1:{}1'.format(last_tab))
    ws.freeze_panes = ws['a2']
    
    return ws


def format_status_table(ws, last_tab):
    default_column_width = 37

    ws.merge_cells('b1:c1')
    ws.merge_cells('e1:f1')
    if str(ws['b1'].value) == 'None':
        ws['b1'].value = ws['a1'].value
        ws['b1'].font = Font(bold=True, name='Dialog.bold')
        ws['a1'].value = ""
    set_header(ws, 'A1:{}2'.format(last_tab))
    ws.freeze_panes = ws['a3']
    for i in list(string.ascii_lowercase):
        ws.column_dimensions[i].width = default_column_width
        if i == "a" : ws.column_dimensions[i].width = "12"
        if "IDENTIFIER" in str(ws['{}2'.format(i)].value) or ws['{}1'.format(i)].value == "to":
            ws.column_dimensions[i].width = "4"
            ws['{}2'.format(i)].alignment = Alignment(horizontal='left')
            ws['{}1'.format(i)].alignment = Alignment(horizontal='left')
    return ws


# def menu():
#     title = 'Please choose your favorite programming language: '
#     options = ['Template 1', 'Template 2']
#     option, index = pick(options, title)

    return index

def find_last_tab_2(ws):
    col_range = list(string.ascii_lowercase)
    # print(str(ws.title),"------- title ----------------")
    for c in col_range:
        # print(ws['{}2'.format(c)].value)
        if "Document Status" in str(ws['{}1'.format(c)].value) and c != "a" :
            # print("found",c, ws.title)
            return next_alpha(c)
    return "f"


def find_last_tab(ws):
    col_range = list(string.ascii_lowercase)
    # print(str(ws.title),"------- title ----------------")
    for c in col_range:
        # print(ws['{}1'.format(c)].value)
        if  "CERTEX" in str(ws['{}1'.format(c)].value):
            # print("found",c, ws.title)
            return c
    return "d"


def main():
    done_folder = "done"
    work_folder = "work"
    full_work_folder = os.path.join(os.path.dirname(os.path.realpath(__file__)), work_folder, "")
    dir_list = []

    for file in  os.listdir(full_work_folder):
        if file.endswith(".xlsx"):
            dir_list.append(file)

    print("Formatting all {} .xlsx files found in {}".format(len(dir_list), full_work_folder))   

    for filename in tqdm(dir_list, desc ="Work done: "):
        if "Transformation Table" in filename and "Status" not in filename:
            wb = openpyxl.load_workbook(os.path.join(full_work_folder, filename))
            for ws_name in wb.sheetnames:
                ws = wb[ws_name]
                ws = format_first_type(ws)
            wb.save(os.path.join(os.path.dirname(os.path.realpath(__file__)), done_folder, filename))
        elif "Information Result" in filename:
            wb = openpyxl.load_workbook(os.path.join(full_work_folder, filename))
            for ws_name in wb.sheetnames:
                if "Recap" not in ws_name:
                    # last_tab = find_last_tab(ws)
                    # print(last_tab)
                    ws = wb[ws_name]
                    ws = format_information_result(ws, find_last_tab(ws))
            wb.save(os.path.join(os.path.dirname(os.path.realpath(__file__)), done_folder, filename))
        elif "Status Transformation Table" in filename:
            wb = openpyxl.load_workbook(os.path.join(full_work_folder, filename))
            for ws_name in wb.sheetnames:
                if "Statuses" not in ws_name:
                    # last_tab = find_last_tab(ws)
                    # print(last_tab)
                    ws = wb[ws_name]
                    ws = format_status_table(ws, find_last_tab_2(ws))
            wb.save(os.path.join(os.path.dirname(os.path.realpath(__file__)), done_folder, filename))

 
if __name__ == '__main__':
    main()