#Show what went wrong without quickly closing program
def show_exception(exc_type, exc_value, tb):
	import traceback
	traceback.print_exception(exc_type, exc_value, tb)
	raw_input("press any key to exit.")
	sys.exit(-1)

import sys
sys.excepthook = show_exception
import os
import shutil 
import time

starting_rfi = input("Please enter starting rfi number: ")
ending_rfi = input("Please enter ending rfi number: ")

for rfi in range(starting_rfi, ending_rfi + 1):
	current_rfi = str(rfi)
	print("RFI: " + current_rfi)
#	print("Source: " + source)	#Show user source 
	if (os.path.isfile("C:/Users/gsavage/Downloads/RFI_0" + current_rfi + "-00_Response.pdf")):
		print ("There is an attachment: " + str(os.path.isfile("C:/Users/gsavage/Downloads/RFI_0" + current_rfi + "-00_Response.pdf")))
		source = "C:/Users/gsavage/Downloads/RFI_0" + current_rfi + "-00_Response.pdf"
		destination = os.path.join("C:/Users/gsavage/Desktop/RFI_Downloads/", current_rfi)
#		print("Destination: " + destination)
		try:
			os.makedirs(destination)
		except OSError:
			print ("Folder " + current_rfi + " Already Exists")
		try:
			shutil.move(source, destination)
		except shutil.Error:
			print ("Response " + current_rfi + " Already Stored")

	elif (os.path.isfile("C:/Users/gsavage/Downloads/RFI_0" + current_rfi + "_00_Response.pdf")):
		print ("There is an attachment: " + str(os.path.isfile("C:/Users/gsavage/Downloads/RFI_0" + current_rfi + "_00_Response.pdf")))
		source = "C:/Users/gsavage/Downloads/RFI_0" + current_rfi + "_00_Response.pdf"
		destination = os.path.join("C:/Users/gsavage/Desktop/RFI_Downloads/", current_rfi)
#		print("Destination: " + destination)
		try:
			os.makedirs(destination)
		except OSError:
			print ("Folder " + current_rfi + " Already Exists")
		try:
			shutil.move(source, destination)
		except shutil.Error:
			print ("Response " + current_rfi + " Already Stored")


	elif (os.path.isfile("C:/Users/gsavage/Downloads/RFI_" + current_rfi + "_PJF_Response.pdf")):
		print ("There is an attachment: " + str(os.path.isfile("C:/Users/gsavage/Downloads/RFI_" + current_rfi + "_PJF_Response.pdf")))
		source = "C:/Users/gsavage/Downloads/RFI_" + current_rfi + "_PJF_Response.pdf" 
		destination = os.path.join("C:/Users/gsavage/Desktop/RFI_Downloads/", current_rfi)
#		print("Destination: " + destination)
		try:
			os.makedirs(destination)
		except OSError:
			print ("Folder " + current_rfi + " Already Exists")
		try:
			shutil.move(source, destination)
		except shutil.Error:
			print ("Response " + current_rfi + " Already Stored")

	elif (os.path.isfile("C:/Users/gsavage/Downloads/RFI-" + current_rfi + "_PJF_Response.pdf")):
		print ("There is an attachment: " + str(os.path.isfile("C:/Users/gsavage/Downloads/RFI-" + current_rfi + "_PJF_Response.pdf")))
		source = "C:/Users/gsavage/Downloads/RFI-" + current_rfi + "_PJF_Response.pdf" 
		destination = os.path.join("C:/Users/gsavage/Desktop/RFI_Downloads/", current_rfi)
#		print("Destination: " + destination)
		try:
			os.makedirs(destination)
		except OSError:
			print ("Folder " + current_rfi + " Already Exists")
		try:
			shutil.move(source, destination)
		except shutil.Error:
			print ("Response " + current_rfi + " Already Stored")
	
	elif (os.path.isfile("C:/Users/gsavage/Downloads/RFI_0" + current_rfi + "_PJF_Response.pdf")):
		print ("There is an attachment: " + str(os.path.isfile("C:/Users/gsavage/Downloads/RFI_0" + current_rfi + "_PJF_Response.pdf")))
		source ="C:/Users/gsavage/Downloads/RFI_0" + current_rfi + "_PJF_Response.pdf" 
		destination = os.path.join("C:/Users/gsavage/Desktop/RFI_Downloads/", current_rfi)
#		print("Destination: " + destination)
		try:
			os.makedirs(destination)
		except OSError:
			print ("Folder " + current_rfi + " Already Exists")
		try:
			shutil.move(source, destination)
		except shutil.Error:
			print ("Response " + current_rfi + " Already Stored")
	
	elif (os.path.isfile("C:/Users/gsavage/Downloads/RFI_" + current_rfi + "_Sketch-PJF_Response.pdf")):
		print ("There is an attachment: " + str(os.path.isfile("C:/Users/gsavage/Downloads/RFI_" + current_rfi + "_Sketch-PJF_Response.pdf")))
		source ="C:/Users/gsavage/Downloads/RFI_" + current_rfi + "_Sketch-PJF_Response.pdf" 
		destination = os.path.join("C:/Users/gsavage/Desktop/RFI_Downloads/", current_rfi)
#		print("Destination: " + destination)
		try:
			os.makedirs(destination)
		except OSError:
			print ("Folder " + current_rfi + " Already Exists")
		try:
			shutil.move(source, destination)
		except shutil.Error:
			print ("Response " + current_rfi + " Already Stored")

	elif (os.path.isfile("C:/Users/gsavage/Downloads/RFI_" + current_rfi + "_HKS_Response.pdf")):
		print ("There is an attachment: " + str(os.path.isfile("C:/Users/gsavage/Downloads/RFI_" + current_rfi + "_HKS_Response.pdf")))
		source ="C:/Users/gsavage/Downloads/RFI_" + current_rfi + "_HKS_Response.pdf" 
		destination = os.path.join("C:/Users/gsavage/Desktop/RFI_Downloads/", current_rfi)
#		print("Destination: " + destination)
		try:
			os.makedirs(destination)
		except OSError:
			print ("Folder " + current_rfi + " Already Exists")
		try:
			shutil.move(source, destination)
		except shutil.Error:
			print ("Response " + current_rfi + " Already Stored")
	
	elif (os.path.isfile("C:/Users/gsavage/Downloads/RFI_" + current_rfi + "_Response.pdf")):
		print ("There is an attachment: " + str(os.path.isfile("C:/Users/gsavage/Downloads/RFI_" + current_rfi + "_Response.pdf")))
		source = "C:/Users/gsavage/Downloads/RFI_" + current_rfi + "_Response.pdf"
		destination = os.path.join("C:/Users/gsavage/Desktop/RFI_Downloads/", current_rfi)
#		print("Destination: " + destination)
		try:
			os.makedirs(destination)
		except OSError:
			print ("Folder " + current_rfi + " Already Exists")
		try:
			shutil.move(source, destination)
		except shutil.Error:
			print ("Response " + current_rfi + " Already Stored")	

	elif (os.path.isfile("C:/Users/gsavage/Downloads/RFI-" + current_rfi + "_Response.pdf")):
		print ("There is an attachment: " + str(os.path.isfile("C:/Users/gsavage/Downloads/RFI-" + current_rfi + "_Response.pdf")))
		source = "C:/Users/gsavage/Downloads/RFI-" + current_rfi + "_Response.pdf"
		destination = os.path.join("C:/Users/gsavage/Desktop/RFI_Downloads/", current_rfi)
#		print("Destination: " + destination)
		try:
			os.makedirs(destination)
		except OSError:
			print ("Folder " + current_rfi + " Already Exists")
		try:
			shutil.move(source, destination)
		except shutil.Error:
			print ("Response " + current_rfi + " Already Stored")	

	elif (os.path.isfile("C:/Users/gsavage/Downloads/RFI_0" + current_rfi + "_Response.pdf")):
		print ("There is an attachment: " + str(os.path.isfile("C:/Users/gsavage/Downloads/RFI_0" + current_rfi + "_Response.pdf")))
		source = "C:/Users/gsavage/Downloads/RFI_0" + current_rfi + "_Response.pdf"
		destination = os.path.join("C:/Users/gsavage/Desktop/RFI_Downloads/", current_rfi)
#		print("Destination: " + destination)
		try:
			os.makedirs(destination)
		except OSError:
			print ("Folder " + current_rfi + " Already Exists")
		try:
			shutil.move(source, destination)
		except shutil.Error:
			print ("Response " + current_rfi + " Already Stored")	
	
	else:
		print ("There is an attachment: False")
	print("")
#	time.sleep(0.1)
	
raw_input("Press any key to exit.")
sys.exit(-1)
