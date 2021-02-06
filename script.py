import qrcode
import openpyxl
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

		phoneNumber = wb_obj.active["C"+str(counter)].value

		if phoneNumber[0]=="0":
			phoneNumber = "996"+phoneNumber[1:]

		if phoneNumber[0:5] == "99655" or phoneNumber[0:5] == "99657":
			print("MEGACOM!!!", phoneNumber)
			counter+=1
			continue
		if phoneNumber[0:6] == "996755" or phoneNumber[0:6] =="996999" or phoneNumber[0:6] == "996998" or phoneNumber[0:6] == "996995" or phoneNumber[0:6] =="996990" or phoneNumber[0:6] == "996997":
			print("MEGACOM!!!", phoneNumber)
			counter+=1
			continue
		if phoneNumber[0:7] == "9968800" or phoneNumber[0:7] == "9968801" or phoneNumber[0:7] == "9968802" or phoneNumber[0:7] == "9968808" or phoneNumber[0:7] == "9968809":
			print("MEGACOM!!!", phoneNumber)
			counter+=1
			continue



		data = ("https://balance.kg/pay/"+
			wb_obj.active["D"+str(counter)].value+
			wb_obj.active["B2"].value+camel_Snake_Case+
			wb_obj.active["C2"].value+phoneNumber
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
