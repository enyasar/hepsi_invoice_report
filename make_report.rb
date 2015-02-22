require 'xlsx2json'
require 'json'
require 'writeexcel'

json_path = 'hepsi.json'
excel_path = 'hepsi.xlsx'
sheet_number = 0 # sheet number start from 0
header_row_number = 1 # row number of the header row which contains column names. 
# Rows before this number get ignored. 
# Row numbers start from 1 based on Excel conventions.

Xlsx2json::Transformer.execute excel_path, sheet_number, json_path, header_row_number: header_row_number

#create a new excel file
workbook = WriteExcel.new('hepsi_new.xls')
worksheet = workbook.add_worksheet
title_format = workbook.add_format
title_format.set_bold
title_format.set_color('red')
sub_total_format = workbook.add_format
sub_total_format.set_bold
sub_total_format.set_color('red')

worksheet.write('A1', 'KODU', title_format)
worksheet.write('B1', 'KDV ORAN' , title_format)
worksheet.write('C1', 'TARIH', title_format)
worksheet.write('D1', 'FT NO', title_format)
worksheet.write('E1', 'BRUT TUTAR', title_format)
worksheet.write('F1', 'KDV', title_format)
worksheet.write('G1', 'TUTAR', title_format)
col=0
row=1

#Invoices
invoices = JSON.parse(File.open(json_path).read)
uniq_invoices = invoices.uniq {|s| s['faturano']}

uniq_invoices.each do |invoice|
	selected_invoices = invoices.select {|invoice_number| invoice_number['faturano'] == invoice['faturano']}
	selected_invoices.sort_by! {|tax| tax['kdvoran']}.reverse!
	#tax 8
	sub_brut_total = 0.0
	sub_tax_total = 0.0
	sub_total = 0.0
	tax_1_invoices = selected_invoices.select {|tax_number| tax_number['kdvoran'] == '8'}
	if tax_1_invoices.count > 0
		tax_1_invoices.each do |tax_1_invoice|
			worksheet.write(row, 0, tax_1_invoice['stok_kodu'])
			worksheet.write(row, 1, tax_1_invoice['kdvoran'])
			worksheet.write(row, 2, tax_1_invoice['tarih'])
			worksheet.write(row, 3, tax_1_invoice['faturano'])
			worksheet.write(row, 4, tax_1_invoice['bruttutar'])
			worksheet.write(row, 5, tax_1_invoice['kdv'])
			worksheet.write(row, 6, tax_1_invoice['tutar'])
			row += 1
			sub_brut_total += tax_1_invoice['bruttutar'].to_f
			sub_tax_total += tax_1_invoice['kdv'].to_f
			sub_total += tax_1_invoice['tutar'].to_f
		end
		#tax
			worksheet.write(row, 4, sub_brut_total, sub_total_format)
			worksheet.write(row, 5, sub_tax_total, sub_total_format)
			worksheet.write(row, 6, sub_total, sub_total_format)
		row += 1
	end

	#tax 18
	sub_brut_total = 0.0
	sub_tax_total = 0.0
	sub_total = 0.0
	tax_2_invoices = selected_invoices.select {|tax_number| tax_number['kdvoran'] == '18' and tax_number['stok_kodu'] != 'KRG01' and tax_number['stok_kodu'] != 'KPD01'}
	if tax_2_invoices.count > 0
		tax_2_invoices.each do |tax_2_invoice|
			worksheet.write(row, 0, tax_2_invoice['stok_kodu'])
			worksheet.write(row, 1, tax_2_invoice['kdvoran'])
			worksheet.write(row, 2, tax_2_invoice['tarih'])
			worksheet.write(row, 3, tax_2_invoice['faturano'])
			worksheet.write(row, 4, tax_2_invoice['bruttutar'])
			worksheet.write(row, 5, tax_2_invoice['kdv'])
			worksheet.write(row, 6, tax_2_invoice['tutar'])
			row += 1
			sub_brut_total += tax_2_invoice['bruttutar'].to_f
			sub_tax_total += tax_2_invoice['kdv'].to_f
			sub_total += tax_2_invoice['tutar'].to_f
		end
		#tax
			worksheet.write(row, 4, sub_brut_total.round(4), sub_total_format)
			worksheet.write(row, 5, sub_tax_total, sub_total_format)
			worksheet.write(row, 6, sub_total, sub_total_format)
		row += 1
	end

	#tax KRG01 and KPD01
	sub_brut_total = 0.0
	sub_tax_total = 0.0
	sub_total = 0.0
	tax_3_invoices = selected_invoices.select {|tax_number| tax_number['stok_kodu'] == 'KRG01' or tax_number['stok_kodu'] == 'KPD01'}
	if tax_3_invoices.count > 0
		tax_3_invoices.each do |tax_3_invoice|
			worksheet.write(row, 0, tax_3_invoice['stok_kodu'])
			worksheet.write(row, 1, tax_3_invoice['kdvoran'])
			worksheet.write(row, 2, tax_3_invoice['tarih'])
			worksheet.write(row, 3, tax_3_invoice['faturano'])
			worksheet.write(row, 4, tax_3_invoice['bruttutar'])
			worksheet.write(row, 5, tax_3_invoice['kdv'])
			worksheet.write(row, 6, tax_3_invoice['tutar'])
			row += 1
			sub_brut_total += tax_3_invoice['bruttutar'].to_f
			sub_tax_total += tax_3_invoice['kdv'].to_f
			sub_total += tax_3_invoice['tutar'].to_f
		end
		#tax
			worksheet.write(row, 4, sub_brut_total.round(4), sub_total_format)
			worksheet.write(row, 5, sub_tax_total, sub_total_format)
			worksheet.write(row, 6, sub_total, sub_total_format)
		row += 1
	end
end

#Workbook Close
workbook.close