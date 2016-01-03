#!/usr/bin/env python
# -*- mode: Python; coding: latin-1 -*-
# Time-stamp: "2015-01-24 16:09:55 sb"

#  file       convert_colorbrewer_xls.py
#  copyright  (c) Sebastian Blatt 2015

import xlrd

def cell_array_to_number(cells):
  return map(lambda c: c.value, cells)

if __name__ == "__main__":
  book = xlrd.open_workbook('ColorBrewer_all_schemes_RGBonly3.xls')
  sheet = book.sheet_by_index(0)

  max_row = 1690
  color_name_cells = sheet.col_slice(0, 1, max_row)
  color_r_cells = sheet.col_slice(6, 1, max_row)
  color_g_cells = sheet.col_slice(7, 1, max_row)
  color_b_cells = sheet.col_slice(8, 1, max_row)

  current_scheme = ''
  scheme_start_index = 0

  color_schemes = []

  for i in range(len(color_name_cells)):
    c = color_name_cells[i]
    if c.value != u'':
      if i > 0:
        n = i - scheme_start_index
        color_schemes.append(('%s%d' % (current_scheme.encode('ascii'), n),
                              cell_array_to_number(color_r_cells[scheme_start_index:i]),
                              cell_array_to_number(color_g_cells[scheme_start_index:i]),
                              cell_array_to_number(color_b_cells[scheme_start_index:i])
                            ))
      current_scheme = c.value
      scheme_start_index = i

  with open('colorbrewer.asy', 'w') as f:
    for s in color_schemes:
      x = 'pen[] %s = {' % s[0]
      f.write(x)
      for i in range(len(s[1])):
        if i > 0:
          f.write(' '*len(x))
        f.write('rgb(%g, %g, %g)' % (s[1][i] / 255.0,
                                     s[2][i] / 255.0,
                                     s[3][i] / 255.0));
        if i < len(s[1])-1:
          f.write(",\n")
      f.write("};\n\n")





# convert_colorbrewer_xls.py ends here
