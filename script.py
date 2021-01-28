import qrcode
import openpyxl
from pathlib import Path
import re
import os
import zipfile
import sys



def main():
	if len(sys.argv) > 1:
		docName = sys.argv[1]
	else:
		docName = "data.xlsx"
	try:
		wb_obj = openpyxl.load_workbook(docName) # add exel file
	except:
		return "error document"+docName

	counter = 3;

	folderName = wb_obj.active["F"+str(counter)].value


# create folder ###############################################
	try:
		os.mkdir(folderName)

	except OSError:
		return "Creation of the directory %s failed" % folderName 
		
	else:
	    print ("Successfully created the directory %s " % folderName)



	while(wb_obj.active["B"+str(counter)].value != None):
	# Formation data for link ##################################	
		mass = wb_obj.active["D"+str(counter)].value.split("_")
		camel_Snake_Case =""


		for i in mass:
			camel_Snake_Case += i[:1].upper()+i[1:]+"_"
		camel_Snake_Case += wb_obj.active["B"+str(counter)].value


	# Generate output link #####################################
		data = ("https://balance.kg/pay/"+
			wb_obj.active["D"+str(counter)].value+
			wb_obj.active["B2"].value+camel_Snake_Case+
			wb_obj.active["C2"].value+wb_obj.active["C"+str(counter)].value
			)


	# QR image generator #######################################
		qr = qrcode.QRCode(
		    version=1,
		    box_size=10,
		    border=1)

		qr.add_data(data)
		qr.make(fit=True)
		img = qr.make_image(fill='black')
		img.save(folderName+"/"+camel_Snake_Case +  '.png')
	

	# Create image in Zip document
		zip_file = zipfile.ZipFile(folderName + '.zip', 'a')
		zip_file.write(folderName+"/"+camel_Snake_Case +  '.png', compress_type=zipfile.ZIP_DEFLATED)
		zip_file.close()


	#############################################################

		counter+=1
		
		print(data)

	return("Successfully")

print(main())
