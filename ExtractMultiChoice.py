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
			matchList = process.extract(ws, WScatalog, limit=3)
			new_dict[ws] = matchList

	file = Workbook()
	sheet = file.add_sheet('Map')
	
	row = 0
	for short_WS, matchList in new_dict.items():
		col = 1
		sheet.write(row, 0, short_WS)
		for match in matchList:
			sheet.write(row, col, match[0])			
			col = col + 1
			sheet.write(row, col, str(match[1]))
			col = col + 1
		row = row + 1 
	file.save('final_choices.xls')
	
	print("Finished!")

if __name__ == "__main__":
	main()
