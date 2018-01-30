import os
import xlrd
import xlwt
from xlutils.copy import copy
from pandas import DataFrame

dir = raw_input("Insert folder path: ")

filename = "compass.csv"
# opening the file with w+ mode truncates the file
f = open(filename, "w+")
f.write(("%s;%s;%s;%s;%s;%s\n") %('Codice Tabella', 'Importo min', 'Importo max', 'Durata min', 'Durata max', 'Tan'))
f.close()

codice_tabella = []
descrizione = []

for entry in os.listdir(dir):
	if entry.endswith(".xlsx"):

		input = os.path.join(dir, entry)
		code = "compass_" + (input.split("\\")[1]).replace(".xlsx","").replace(" ","_").lower()
		book = xlrd.open_workbook(input)

		codice = list()

		i_ = list()
		i_max = list()
		d_ = list()
		d_max = list()
		t_ = list()

		aim_dur = 0
		aim_tan = 0

		def imp_max(imp_min, length, index2):
			current_value = imp_min[index2]
			return_value = current_value
			for row in range (index2, len(imp_min)):
				if current_value != imp_min[row]:
					return imp_min[row]
					
			return return_value
			
		def dur_max(dur_min, length, index2):
			current_value = dur_min[index2]
			return_value = current_value
			for row in range (index2, len(dur_min)):
				if current_value < dur_min[row]:
					return dur_min[row]
				elif current_value > dur_min[row]:
					return 121
					
			return return_value

		def getMaxImport(current_value, list):
			return_value = current_value
			for item in list:
				if float(return_value) < float(item):
					return item
			return 50000		
			
		def something(index):
			sheet = book.sheet_by_index(index)
			
			importo = list()
			importo_max = list()
			durata = list()
			durata_max = list()
			tan = list()
			
			total_rows = sheet.nrows
			total_cols = sheet.ncols
			
			for i in range(sheet.nrows):
				row = sheet.row_values(i)
				for j in range(len(row)):
					if row[j] == "Durata del":
						aim_dur = j
					
			for i in range(sheet.nrows):
				row = sheet.row_values(i)
				for j in range(len(row)):
					if row[j] == "TAN":
						aim_tan = j
				

		#Appending Importo
			for row in range (4,total_rows-1):
				for col in range(1):
					if str(sheet.cell(row,col).value) != "":
						lastvalue = str((sheet.cell(row,col).value).replace(".","")).split(",")[0]
						
					importo.append("{:.2f}".format(float(lastvalue)))
			
			for index in range(0, len(importo)):
				i_.append(importo[index])

				
		#Appending Durata			
			for row in range (4,total_rows-1):
				for col in range(aim_dur,aim_dur+1):
					durata.append(str(sheet.cell(row,col).value).split(".")[0])
			d_.append(durata)
			
			
		#Appending Durata Max
			for row in range (4,total_rows-1):
				for col in range(aim_dur,aim_dur+1):
					
					lastvalue = dur_max(durata, len(durata), row - 4)
					
					durata_max.append(int(lastvalue) - 1)
					
			d_max.append(durata_max)

		#Appending TAN			
			for row in range (4,total_rows-1):
				for col in range(aim_tan,aim_tan+1):
					tan.append(str(sheet.cell(row,col).value).split(" ")[0].replace(",","."))
			t_.append(tan)
					
		#Appending Codice			
			for row in range (4,total_rows-1):
				codice.append(code)
			

		for index in range(1,book.nsheets):
			something(index)





		im_final = list()
		im_final_max = list()
		dur_final = list()
		dur_final_max = list()
		tan_final = list()

		for sub in i_:
			#for item in sub:
			im_final.append(sub)
				

				
		for sub in im_final:
			lastvalue = getMaxImport(sub, im_final)
			im_final_max.append(float(lastvalue) - 0.01)
			
			
		for sub in d_:
			for item in sub:
				dur_final.append(item)
				
		for sub in d_max:
			for item in sub:
				dur_final_max.append(item)
				
		for sub in t_:
			for item in sub:
				tan_final.append(item)
				
		try:
			impy = im_final.count(im_final[-1])
			dury = dur_final.count(dur_final[-1])

			for i in range(-impy,-0):
				im_final_max[i] = float(500000.00)

			for i in range(-1,-0):
				dur_final_max[i] = 120
				

			print input, " Completed Successfully!"
			
		except IndexError:
			pass
		
		


		df = DataFrame({'Codice Tabella': codice, 'Importo min': im_final, 'Importo max': im_final_max,'Durata min': dur_final, 'Durata max': dur_final_max, 'Tan': tan_final})

		df = df[['Codice Tabella', 'Importo min', 'Importo max', 'Durata min', 'Durata max', 'Tan']]

		#df.to_excel('Final ' + entry + '.xlsx', sheet_name='sheet1', index=False,)

		df.to_csv('compass.csv',";", mode='a',header=False, index=False,)
		
		codice_tabella.append(code)
		descrizione.append(entry)
				
		continue
	else:
		continue
		
print "Done!"

file = open("codici tabella compass.txt","w")

for x,y in zip(codice_tabella, descrizione):
	file.write('"%s"  per il prodotto  "%s"\n' %(x,y.replace(".xlsx","").title()))
	
file.close()

print "codici tabella compass.txt created successfully\n"

raw_input("Press Enter to exit!")
