import xlrd  
from xlutils.copy import copy

rb1 = xlrd.open_workbook('01-24.xlsx') 			# open file
wb1 = copy(rb1) 					   			# load it to memory
sh1 = wb1.get_sheet(0)				   			# get sheet 1 for write
sh2 = wb1.get_sheet(1)				   			# get sheet 2 for write
sheet1 = rb1.sheet_by_index(0)         			# get sheet 1 for read
sheet2 = rb1.sheet_by_index(1)         			# get sheet 2 for read
counter = 0							   			# hz naxui

for row1 in range(sheet1.nrows):	   			# iter rows of kassa
	cellv1 = ''						   			# var for compare
	for col1 in (3,4):							# getting 2 col value and make comp str
		cell1 = sheet1.cell_value(row1,col1)	#
		cellv1 += str(cell1)					# comp str1

		for row2 in range(sheet2.nrows):		# iter rows of billing
			cellv2 = ''							# var for compare
			for col2 in (5,6):					# getting 2 col value and make comp str
				cell2 = sheet2.cell_value(row2,col2)
				cellv2 += str(cell2)			# comp str2
			if(cellv1==cellv2):					# comparation
				counter+=1						# hz naxui

				print(str(counter) + ".  " + str(row1+1)+'  -  '+str(sheet1.cell_value(row1,0))+" "+str(sheet1.cell_value(row1,1))+' == '+str(sheet2.cell_value(row1,0))+" "+str(sheet2.cell_value(row1,1))+"  -  " + str(row2+1)) #log

				sh1.write(row1, 5, ' '+str(row2+1))		# addding comment
				sh2.write(row2, 7, ' '+str(row1+1))		# addding comment
				
wb1.save('result.xls')							#saving
