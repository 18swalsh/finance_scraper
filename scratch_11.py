#xlsxwriter docucmentation reference

import pandas as pd
import xlsxwriter



#will open an excel file and put 'hello world' in A1 of 'Sheet1'
workbook = xlsxwriter.Workbook('Output.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write(0, 0, 'Hello Excel')
workbook.close()

#will do the same as above, no need for a close() statement

with xlsxwriter.Workbook('hello_world.xlsx') as workbook:
    worksheet = workbook.add_worksheet()
    worksheet.write('A1', 'Hello world')

#uses float() to convert strings to numbers where possible
workbook = xlsxwriter.Workbook(filename, {'strings_to_numbers': True})
                            # other options
                            # (filename, {'strings_to_formulas': False})
                            # (filename, {'strings_to_urls': False})
                            # (filename, {'nan_inf_to_errors': True})
                            # (filename, {'default_date_format': 'dd/mm/yy'})
                            # (filename, {'remove_timezone': True})

#add_worksheet('sheet_name')
#cannot contain ' [ ] : * ? / \ '
#cannot use the same name for multiple worksheets

#add_format()
format1 = workbook.add_format()       # Set properties later.
format2 = workbook.add_format(props)  # Set properties at creation.

format = workbook.add_format()
format.set_bold()
format.set_font_color('red')

format = workbook.add_format({'bold': True, 'font_color': 'red'})

worksheet.write       (0, 0, 'Foo', format)
worksheet.write_string(1, 0, 'Bar', format)
worksheet.write_number(2, 0, 3,     format)
worksheet.write_blank (3, 0, '',    format)

worksheet.set_row(0, 18, format)
worksheet.set_column('A:D', 20, format)

format = workbook.add_format()

format.set_bold()      # Turns bold on.
format.set_bold(True)  # Also turns bold on.
format.set_bold(False)  # Turns bold off.

cell_format.set_font_name('Times New Roman')

# Category	Description	Property	Method Name
# Font	Font type	'font_name'	set_font_name()
#  	Font size	'font_size'	set_font_size()
#  	Font color	'font_color'	set_font_color()
#  	Bold	'bold'	set_bold()
#  	Italic	'italic'	set_italic()
#  	Underline	'underline'	set_underline()
#  	Strikeout	'font_strikeout'	set_font_strikeout()
#  	Super/Subscript	'font_script'	set_font_script()
# Number	Numeric format	'num_format'	set_num_format()
# Protection	Lock cells	'locked'	set_locked()
#  	Hide formulas	'hidden'	set_hidden()
# Alignment	Horizontal align	'align'	set_align()
#  	Vertical align	'valign'	set_align()
#  	Rotation	'rotation'	set_rotation()
#  	Text wrap	'text_wrap'	set_text_wrap()
#  	Justify last	'text_justlast'	set_text_justlast()
#  	Center across	'center_across'	set_center_across()
#  	Indentation	'indent'	set_indent()
#  	Shrink to fit	'shrink'	set_shrink()
# Pattern	Cell pattern	'pattern'	set_pattern()
#  	Background color	'bg_color'	set_bg_color()
#  	Foreground color	'fg_color'	set_fg_color()
# Border	Cell border	'border'	set_border()
#  	Bottom border	'bottom'	set_bottom()
#  	Top border	'top'	set_top()
#  	Left border	'left'	set_left()
#  	Right border	'right'	set_right()
#  	Border color	'border_color'	set_border_color()
#  	Bottom color	'bottom_color'	set_bottom_color()
#  	Top color	'top_color'	set_top_color()
#  	Left color	'left_color'	set_left_color()
#  	Right color	'right_color'	set_right_color()

set_num_format('0 "dollar and" .00 "cents"')

# Index	Index	Format String
# 0	0x00	General
# 1	0x01	0
# 2	0x02	0.00
# 3	0x03	#,##0
# 4	0x04	#,##0.00
# 5	0x05	($#,##0_);($#,##0)
# 6	0x06	($#,##0_);[Red]($#,##0)
# 7	0x07	($#,##0.00_);($#,##0.00)
# 8	0x08	($#,##0.00_);[Red]($#,##0.00)
# 9	0x09	0%
# 10	0x0a	0.00%
# 11	0x0b	0.00E+00
# 12	0x0c	# ?/?
# 13	0x0d	# ??/??
# 14	0x0e	m/d/yy
# 15	0x0f	d-mmm-yy
# 16	0x10	d-mmm
# 17	0x11	mmm-yy
# 18	0x12	h:mm AM/PM
# 19	0x13	h:mm:ss AM/PM
# 20	0x14	h:mm
# 21	0x15	h:mm:ss
# 22	0x16	m/d/yy h:mm
# ...	...	...
# 37	0x25	(#,##0_);(#,##0)
# 38	0x26	(#,##0_);[Red](#,##0)
# 39	0x27	(#,##0.00_);(#,##0.00)
# 40	0x28	(#,##0.00_);[Red](#,##0.00)
# 41	0x29	_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)
# 42	0x2a	_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)
# 43	0x2b	_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)
# 44	0x2c	_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)
# 45	0x2d	mm:ss
# 46	0x2e	[h]:mm:ss
# 47	0x2f	mm:ss.0
# 48	0x30	##0.0E+0
# 49	0x31	@


worksheets()
for worksheet in workbook.worksheets():
    worksheet.write('A1', 'Hello')

#get_worksheet_by_name() --> returns a worksheet object in the workbook
worksheet = workbook.get_worksheet_by_name('Sheet1')