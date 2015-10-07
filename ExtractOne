#!/usr/bin/env python

import xlrd
from xlwt import Workbook
from fuzzywuzzy import process
from collections import OrderedDict

def main():
	#get file names as inputs here later
	file = xlrd.open_workbook("WSnotof.xls")
	sheet = file.sheets()[0]
	WSnotof = sheet.col_values(2, 1)

	file = xlrd.open_workbook("Catalog.xlsx")
	sheet = file.sheets()[0]
	WScatalog = sheet.col_values(2, 1)

	new_dict = OrderedDict()
	#new_dict = dict()
	for ws in WSnotof:
		if not ws in new_dict:
			match = process.extractOne(ws, WScatalog)
			new_dict[ws] = match

	file = Workbook()
	sheet = file.add_sheet('Map')
	
	row = 0
	for short_WS, WS in new_dict.items():
		sheet.write(row, 0, str(WS[1]))
		sheet.write(row, 1, short_WS)
		sheet.write(row, 2, WS[0])
		row = row + 1 
	file.save('final.xls')
	
	print("Finished!")

if __name__ == "__main__":
	main()
