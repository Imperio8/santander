import os
import xlrd
import xlwt
import operator

dir = raw_input("Insert folder path: ")


for entry in os.listdir(dir):
	if entry.endswith(".xlsx"):
	
		input = os.path.join(dir, entry)
		book = xlrd.open_workbook(input)  
		
		for sheet in book.sheets():  
			total_rows = sheet.nrows
			total_cols = sheet.ncols
			
			durata = []
			tan = []
			coeff= []
			spese = []
			comm = []
			tot = []
			
			for row in range (total_rows):
				for col in range(1):
					asd = sheet.cell(row,col).value
					if type(asd) == float:
						durata.append(asd)
						
			for row in range (total_rows):
				for col in range(2):
					asd = sheet.cell(row,col).value
					if type(asd) == float:
						tan.append(asd)
						
			for row in range (total_rows):
				for col in range(3):
					asd = sheet.cell(row,col).value
					if type(asd) == float:
						coeff.append(asd)
						
			for row in range (total_rows):
				for col in range(4):
					asd = sheet.cell(row,col).value
					if type(asd) == float:
						spese.append(asd)
						
			for row in range (total_rows):
				for col in range(5):
					asd = sheet.cell(row,col).value
					if type(asd) == float:
						comm.append(asd)
						
			for row in range (total_rows):
				for col in range(6):
					asd = sheet.cell(row,col).value
					if type(asd) == float:
						tot.append(asd)
						
			
			
			for q,w,e,r,t,y in zip(durata,tan, coeff, spese, comm, tot):
				print q , '\t\t\t' , w , '\t\t\t' , e , '\t\t\t' , r , '\t\t\t' , t , '\t\t\t' , y
			print "----------------"
			
		