import qrcode
import openpyxl
from pathlib import Path
import re

wb_obj = openpyxl.load_workbook("data.xlsx")

counter = 3;

while(wb_obj.active["B"+str(counter)].value != None):
	mass = wb_obj.active["D"+str(counter)].value.split("_")
	camel_Snake_Case =""


	for i in mass:
		camel_Snake_Case += i[:1].upper()+i[1:]+"_"
	camel_Snake_Case += wb_obj.active["B"+str(counter)].value


	data = ("https://balance.kg/pay/"+
		wb_obj.active["D"+str(counter)].value+
		wb_obj.active["B2"].value+camel_Snake_Case+
		wb_obj.active["C2"].value+wb_obj.active["C"+str(counter)].value
		)

	qr = qrcode.QRCode(
	    version=1,
	    box_size=10,
	    border=1)

	qr.add_data(data)
	qr.make(fit=True)
	img = qr.make_image(fill='black')
	img.save(camel_Snake_Case +  '.png')

	counter+=1
	
	print(data)
	

# https://balance.kg/pay/ServiceCode?requisite=Rakhmanov_Kenzhebaj&notify=996553556411
