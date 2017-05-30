from openpyxl import load_workbook, worksheet

def WriteData(filename, students):
    firstname_col = 2
    lastname_col = 3
    email_col = 4
    score_col = 15
    course_col = 12

    wb = load_workbook(filename = filename)
    sheet = wb.get_sheet_by_name('Attendees')
    i = 0
    for row in sheet.iter_rows():
        i += 1
        #print(row.index())
        print(sheet.cell(row = i, column = firstname_col).value) 
        #for cell in row:
            #print(dir(cell))
            #print(cell.column)
        #c = sheet.cell(row = row, column = firstname_col).value
        #print(c)

# ----- ROW PROPERTIES -----
# ['__add__', '__class__', '__contains__', '__delattr__', '__dir__', '__doc__', '__eq__', '__format__', '__ge__', '__getattribute__', '__g
# etitem__', '__getnewargs__', '__gt__', '__hash__', '__init__', '__init_subclass__', '__iter__', '__le__', '__len__', '__lt__', '__mul__'
# , '__ne__', '__new__', '__reduce__', '__reduce_ex__', '__repr__', '__rmul__', '__setattr__', '__sizeof__', '__str__', '__subclasshook__'
# , 'count', 'index']

# ----- COLUMN PROPERTIES -----
#  'row', 'set_explicit_value', 'style', 'style_id', 'value']
# ['ERROR_CODES', 'TYPE_BOOL', 'TYPE_ERROR', 'TYPE_FORMULA', 'TYPE_FORMULA_CACHE_STRING', 'TYPE_INLINE', 'TYPE_NULL', 'TYPE_NUMERIC', 'TYP
# E_STRING', 'VALID_TYPES', '__class__', '__delattr__', '__dir__', '__doc__', '__eq__', '__format__', '__ge__', '__getattribute__', '__gt_
# _', '__hash__', '__init__', '__init_subclass__', '__le__', '__lt__', '__module__', '__ne__', '__new__', '__reduce__', '__reduce_ex__', '
# __repr__', '__setattr__', '__sizeof__', '__slots__', '__str__', '__subclasshook__', '_bind_value', '_cast_datetime', '_cast_numeric', '_
# cast_percentage', '_cast_time', '_comment', '_hyperlink', '_infer_value', '_style', '_value', 'alignment', 'anchor', 'base_date', 'borde
# r', 'check_error', 'check_string', 'col_idx', 'column', 'comment', 'coordinate', 'data_type', 'encoding', 'fill', 'font', 'guess_types',
#  'has_style', 'hyperlink', 'internal_value', 'is_date', 'number_format', 'offset', 'parent', 'pivotButton', 'protection', 'quotePrefix',
#  'row', 'set_explicit_value', 'style', 'style_id', 'value']

