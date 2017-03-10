# -*- coding: utf-8 -*-
# Copyright (c) 2012-2013 by Pablo Martín <goinnn@gmail.com>
#
# This software is free software: you can redistribute it and/or modify
# it under the terms of the GNU Lesser General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This software is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public License
# along with this software.  If not, see <http://www.gnu.org/licenses/>.

import csv
import collections
import sys
import xlwt

from .base import get_content

PY3 = sys.version_info[0] == 3


# http://stackoverflow.com/a/214657
def hex_to_rgb(value):
    """Return (red, green, blue) for the color given as #rrggbb."""
    value = value.lstrip('#')
    lv = len(value)
    return tuple(int(value[i:i + lv // 3], 16) for i in range(0, lv, lv // 3))


def convert(response, encoding='utf-8', title_sheet='Sheet 1',  content_attr='content', csv_kwargs=None):
    """Replace HttpResponse csv content with excel formatted data using xlwt
    library.
    """
    csv_kwargs = csv_kwargs or {}
    # Styles used in the spreadsheet.  Headings are bold.
    header_font = xlwt.Font()
    header_font.bold = True

    header_style = xlwt.XFStyle()
    header_style.font = header_font

    wb = xlwt.Workbook(encoding=encoding)
    
    # Parse style response and add colors to palette as necessary
    response_style = get_content(response['style'])
    reader_style = csv.reader(response_style, **csv_kwargs)
    custom_colors = []
    custom_colors_index = 8
    for lno, line_style in enumerate(reader_style):
        if lno > 0:
            for cno, utf8_text in enumerate(line_style):
                cell_text = utf8_text
                if not PY3:
                    cell_text = cell_text.decode(encoding)
                if cell_text != u"" and cell_text not in custom_colors:
                    custom_colors.append(cell_text)
                    # Add new color to palette and set RGB color value
                    color_name = u"color_{}".format(cell_text.lstrip('#'))
                    xlwt.add_palette_colour(color_name, custom_colors_index)
                    color_RGB = hex_to_rgb(cell_text) 
                    wb.set_colour_RGB(custom_colors_index, color_RGB[0], color_RGB[1], color_RGB[2])
                    custom_colors_index = custom_colors_index + 1
    
    ws = wb.add_sheet(title_sheet)

    # Cell width information kept for every column, indexed by column number.
    cell_widths = collections.defaultdict(lambda: 0)
    
    response_style = get_content(response['style']) # open again to reset iterator
    reader_style = csv.reader(response_style, **csv_kwargs) # open again to reset iterator
    
    response = response['content']
    reponse_content = get_content(response)    
    reader_content = csv.reader(reponse_content, **csv_kwargs)
    
    for lno, (line_content, line_style) in enumerate(zip(reader_content, reader_style)):
        if lno == 0:
            row_style = header_style
        else:
            row_style = None
        write_row(ws, lno, line_content, line_style, cell_widths, style=row_style, encoding=encoding)
    
    # Roughly autosize output column widths based on maximum column size.
    for col, width in cell_widths.items():
        ws.col(col).width = width
    setattr(response, content_attr, '')
    wb.save(response)
    return response


def write_row(ws, lno, cell_text, cell_style, cell_widths, style=None, encoding='utf-8'):
    """Write row of utf-8 encoded data to worksheet, keeping track of maximum
    column width for each cell.
    """
    import xlwt
    if style is None:
        style = xlwt.Style.default_style
    for cno, (utf8_text, utf8_style) in enumerate(zip(cell_text, cell_style)):
        cell_text = utf8_text
        if not PY3:
            cell_text = cell_text.decode(encoding)
        
        if lno > 0:
            if utf8_style is not "":
                # Override style with background color
                style = xlwt.easyxf('pattern: pattern solid, fore_colour color_{}'.format(utf8_style.lstrip('#')))
            else:
                style = xlwt.Style.default_style
        ws.write(lno, cno, cell_text, style)
        cell_widths[cno] = max(cell_widths[cno],
                               get_xls_col_width(cell_text, style))


# A reasonable approximation for column width is based off zero character in
# default font.  Without knowing exact font details it's impossible to
# determine exact auto width.
# http://stackoverflow.com/questions/3154270/python-xlwt-adjusting-column-widths?lq=1
def get_xls_col_width(text, style):
    return int((1 + len(text)) * 256)
