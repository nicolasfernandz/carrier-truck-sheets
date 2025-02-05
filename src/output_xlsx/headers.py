def get_cell_format(workbook, color='orange', top=1, left=1, right=1, bottom=1, size=11, setnumformat=True, numformat='$ #,##0', center=False, align_left=False, align_right=False, liquid_de="EDISON"):
    if color == 'orange':
        format_color = '#FDE9D9'
        if liquid_de == "EDU":
            format_color = '#E6F4EA'
    elif color == 'grey':
        format_color = '#d4d5d3'
    elif color == 'header':
        format_color = '#d9d9d9'
    else:
        format_color = '#FCD5B4'
        if liquid_de == "EDU":
            format_color = '#C4DECB'

    if setnumformat:
        # Add a header format.
        format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'fg_color': format_color,
            'top':top,
            'left':left,
            'right': right,
            'bottom':bottom,
            'num_format': numformat}) #'$ #,##0'
    else:
        # Add a header format.
        format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'fg_color': format_color,
            'top':top,
            'left':left,
            'right': right,
            'bottom':bottom})

    if center:
        format.set_align('center')
        format.set_align('vcenter')

    if align_left:
        format.set_align('left')
        format.set_align('vleft')

    if align_right:
        format.set_align('right')
        format.set_align('vright')

    format.set_font_size(size)
    return format

def get_header_cell_format(workbook):
    format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'align' : 'center',
        'fg_color': '#d9d9d9'})

    return format

def get_header_cell_format_for_merged_cols(workbook):
    format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'align' : 'center',
        'fg_color': '#d9d9d9',
        'top':0,
        'left':0,
        'right': 2,
        'bottom':2})

    return format

def get_header_cell_format_fo_last_col(workbook):
    format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'align' : 'center',
        'fg_color': '#d9d9d9',
        'top':0,
        'left':0,
        'right': 2,
        'bottom':0})

    return format
