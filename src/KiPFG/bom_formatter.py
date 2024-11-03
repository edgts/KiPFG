import xlsxwriter

workbook = xlsxwriter.Workbook('test.xlsx')
worksheet = workbook.add_worksheet()

digits = 1
continuous = 1
dotted = 7

ttlbrd = workbook.add_format()
ttlbrd.set_num_format(digits)
ttlbrd.set_top(continuous)
ttlbrd.set_bottom(continuous)
ttlbrd.set_left(dotted)
ttlbrd.set_right(dotted)
ttlbrd.set_align('vcenter')

ttllft = workbook.add_format()
ttllft.set_num_format(digits)
ttllft.set_top(continuous)
ttllft.set_bottom(continuous)
ttllft.set_left(continuous)
ttllft.set_right(dotted)
ttllft.set_align('vcenter')

ttlrgt = workbook.add_format()
ttlrgt.set_num_format(digits)
ttlrgt.set_top(continuous)
ttlrgt.set_bottom(continuous)
ttlrgt.set_left(dotted)
ttlrgt.set_right(continuous)
ttlrgt.set_align('vcenter')

dotbrd = workbook.add_format()
dotbrd.set_num_format(digits)
dotbrd.set_bottom(dotted)
dotbrd.set_left(dotted)
dotbrd.set_right(dotted)
dotbrd.set_text_wrap()

numfmt = workbook.add_format()
numfmt.set_num_format(digits)

cntfmt = workbook.add_format()
cntfmt.set_num_format(digits)
cntfmt.set_top(continuous)
cntfmt.set_bottom(continuous)
cntfmt.set_left(continuous)
cntfmt.set_right(continuous)

emptyfmt = workbook.add_format()
emptyfmt.set_num_format(digits)
emptyfmt.set_bottom(dotted)
emptyfmt.set_left(dotted)
emptyfmt.set_right(dotted)
emptyfmt.set_text_wrap()
emptyfmt.set_align('vcenter')
emptyfmt.set_text_wrap()
emptyfmt.set_bg_color('yellow')

# configure view
worksheet.set_paper(9)  # A4
worksheet.set_landscape()
worksheet.set_header("&LDatei: &F&RSeite &P/&N")
worksheet.freeze_panes(5, 0)
worksheet.repeat_rows(0, 4)
# worksheet.set_default_row(30, hide_unused_rows=False)
worksheet.hide_gridlines(1)
worksheet.fit_to_pages(1, 0)

# write out title block
row = 0
worksheet.set_row(row, 16)  # set row height
worksheet.write(row, 0, r"Gesellschaft für Test Systeme mbH")
worksheet.write(row, 4, r"Teltower Damm 276")
worksheet.write(row, 5, r"D-14167 Berlin")
worksheet.write(row, 6, r"Tel.:+4930/845723-0")
worksheet.write(row, 7, r"Fax.:+4930/845723-23")

row = 2
worksheet.set_row(row, 16)  # set row height
worksheet.write(row, 0, r"Bestelliste Project 1, Project 2 Revision: 1")
worksheet.write(row, 4, "Erstellt: 2024-11-02")
worksheet.write(row, 6, "Project Description; revised: 2024-06-22")

row = 3
worksheet.set_row(row, 16)  # set row height
worksheet.write(row, 0, r"Anzahl der zu fertigenden Geräte eintragen: ==>")
worksheet.write(row, 4, 1, cntfmt)

row = 4
worksheet.set_column(0,  1,  5.00)
worksheet.write(row, 0, r"Lfn", ttllft)
worksheet.write(row, 1, r"Stück", ttlbrd)
worksheet.set_column(2,  2, 16.43)
worksheet.write(row, 2, r"Referenz", ttlbrd)
worksheet.set_column(3,  3, 20.71)
worksheet.write(row, 3, r"Wert", ttlbrd)
worksheet.set_column(4,  4, 25.00)
worksheet.write(row, 4, r"Bezeichnung", ttlbrd)
worksheet.set_column(5,  6, 27.86)
worksheet.write(row, 5, r"Toleranz/Spannung", ttlbrd)
worksheet.write(row, 6, r"Bauform", ttlbrd)
worksheet.set_column(7,  7, 13.57)
worksheet.write(row, 7, r"Hersteller", ttlbrd)
worksheet.set_column(8,  11, 20.71)
worksheet.write(row, 8, r"Bestellnummer", ttlbrd)
worksheet.write(row, 9, r"Alternative BE", ttlbrd)
worksheet.write(row, 10, r"wh Artikelnummer", ttlbrd)
worksheet.write(row, 11, r"Bemerkung", ttlbrd)

worksheet.set_column(12, 13,  5.00)
worksheet.write(row, 12, r"Soll", ttlbrd)
worksheet.write(row, 13, r"Ist", ttlbrd)
worksheet.write(row, 14, r"Variante", ttlrgt)

workbook.close()
