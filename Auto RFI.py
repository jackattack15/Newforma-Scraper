from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import xlwt
import sys
import codecs
import string
import getpass

#This keeps the program open so that you can see what the problem is that you ran into.  Otherwise it would close too fast.
def show_exception(exc_type, exc_value, tb):
	import traceback
	traceback.print_exception(exc_type, exc_value, tb)
	raw_input("press any key to exit.")
	sys.exit(-1)

sys.excepthook = show_exception

if sys.stdout.encoding != 'cp850':
	sys.stdout = codecs.getwriter('cp850')(sys.stdout, 'strict')
if sys.stderr.encoding != 'cp850':
	sys.stderr = codecs.getwriter('cp850')(sys.stderr, 'strict')

#Get User's Newforma Credentials
usrname = raw_input('Newforma Username: ')
psswd = getpass.getpass(prompt = 'Newforma Password: ')
print("")

#This makes it so you don't have to type in your entire email address each time.
usrname = usrname + '@brasfieldgorrie.com'

#We need to know which RFIs to update, so ask the user
start_rfi = input('Please enter starting RFI #: ')
end_rfi = input('Please enter ending RFI #: ')
print("")

#Next log into Newforma
driver = webdriver.Chrome('C:\Python27\Lib\site-packages\selenium\webdriver\chrome\chromedriver.exe')
driver.get("https://projectcloud.newforma.com/")
login_email = driver.find_element_by_name('email')
login_email.send_keys(usrname)
login_pswd = driver.find_element_by_name('pass')
login_pswd.send_keys(psswd)
login_pswd.send_keys(Keys.ENTER)
time.sleep(0.2)
driver.get("https://projectcloud.newforma.com/rfi/dashboard.php")
rfi_log = open('RFI' + str(start_rfi) + str(end_rfi), 'w')

#initiate Excel Workbook
book = xlwt.Workbook()
sheet1 = book.add_sheet('Sheet 1', cell_overwrite_ok=True)

#This sets up the headers of the Excel Sheet.
style = xlwt.easyxf('font: bold 1')
sheet1.write(0, 0 , 'Project RFI Number', style)
sheet1.write(0, 1 , 'Subject', style)
sheet1.write(0, 2 , 'Date Created', style)
sheet1.write(0, 3 , 'Date Required', style)
sheet1.write(0, 4 , 'Date Answered', style)
sheet1.write(0, 5 , 'Authored By', style)
sheet1.write(0, 6 , 'Answer Company', style)
sheet1.write(0, 7 , 'Answered By', style)
sheet1.write(0, 8 , 'Question', style)
sheet1.write(0, 9 , 'Answer', style)
sheet1.write(0, 10, 'Official', style)
sheet1.write(0, 11, 'Closed', style)
#sheet1.write(0, 8, 'Answer 2', style)

#Checks to see if what we are looking for exists on this page by searching for given xpath.
def check_exists_by_xpath(xpath):
	try:
		driver.find_element_by_xpath(xpath).text
	except NoSuchElementException:
		return False
	return True

#Checks to see if what we are looking for exists on this page by searching all the links on the page for the 
# first instance of the given text
def check_exists_by_link(link):
	try:
		driver.find_element_by_partial_link_text(link)
	except NoSuchElementException:
		return False
	return True

#Uses Newforma's "Search" bar to search for the given RFI number.	
def pullup_rfi(rfi, row):
	search = driver.find_element_by_name('search_text')
	search.send_keys(rfi)			#Types the RFI # into the search bar
	search.send_keys(Keys.RETURN)	# then hits Return
	search_rfi = str(rfi)
	print search_rfi

	#looks to see if the RFI we searched for actually exists.
	if check_exists_by_link('RFI-0' + search_rfi):	
		current_rfi = driver.find_element_by_partial_link_text('RFI-0' + search_rfi)

		time.sleep(1)
		current_rfi.click()		#If it does, we open it up by clicking the link.
	
	sheet1.write(row, 0, search_rfi)	#Writes the RFI Number onto the Excel Sheet

#Checks to see if the RFI has a 'Response' attachment.  If it does, it downloads it.
def download_attachments():
	if check_exists_by_link('Response'):
		download = driver.find_element_by_partial_link_text('Response')
		download.click()
		print("Attachment Downloaded")	#Gives the user feedback if there was an attachment
	else:
		print("No Attachment")			#Tells user if there wasn't a response attachment
	
#Gets the title of the RFI
def get_title(row):
	if check_exists_by_xpath("//table[1]/tbody/tr[4]/td[2]"):
		title = driver.find_element_by_xpath("//table[1]/tbody/tr[4]/td[2]").text
		sheet1.write(row, 1, title)		#Writes it to Excel Sheet
		print(title)					#Gives info to the User
		
#Finds whoever wrote the RFI and then deletes the stuff that isn't their name.
def get_projAdmin(row):
	if check_exists_by_xpath("//table[3]/tbody/tr[3]/td[2]"):
		proj_admin = driver.find_element_by_xpath("//table[3]/tbody/tr[3]/td[2]").text
		proj_admin = proj_admin.replace("Project Admin (Contractor) - ","")
		sheet1.write(row, 5, proj_admin)		#Writes it to Excel Sheet
		print("Project Admin:   " + proj_admin)	#Gives info to the user
		
#Gets the date that the RFI was submitted
def get_dateSub(row):
	if check_exists_by_xpath("//table[1]/tbody/tr[7]/td[2]"):
		date_sub = driver.find_element_by_xpath("//table[1]/tbody/tr[7]/td[2]").text
		sheet1.write(row, 2, date_sub)			#Writes it to Excel Sheet
		print("Date Submitted:  " + date_sub) 	#Gives info to the user
		
#Gets the date that the RFI needs to be returned by
def get_dateDue(row):
	if check_exists_by_xpath("//table[1]/tbody/tr[9]/td[2]"):
		due_date = driver.find_element_by_xpath("//table[1]/tbody/tr[9]/td[2]").text
		sheet1.write(row, 3, due_date)			#Writes it to Excel Sheet
		print("Date Required:   " + due_date)	#Gives user info
		
#Gets the date that the RFI actually was returned to us.
def get_dateRet(row):
	if check_exists_by_xpath("//table[1]/tbody/tr[8]/td[2]"):
		date_ret = driver.find_element_by_xpath("//table[1]/tbody/tr[8]/td[2]").text
		sheet1.write(row, 4, date_ret)			#Writes it to Excel Sheet
		print("Date Returned:   " + date_ret)	#Gives the user info

#Gets the question that the RFI is asking.  Doesn't print to User just to save space in the program window.
def get_question(row):
	if check_exists_by_xpath("//table[1]/tbody/tr[10]/td[2]"):
		question = driver.find_element_by_xpath("//table[1]/tbody/tr[10]/td[2]").text
		sheet1.write(row, 8, question)	#Writes the question to the Excel Sheet
		
#Gets the answer to the RFI's question.  Doesn't print to User just to save space in program window.
#There are often several responses, with each newer on being more vague than the last, so we check each response and record it.
def get_answer(row):
	if check_exists_by_xpath("//table[4]/tbody/tr[7]/td[2]/b/div"):
		answer =  driver.find_element_by_xpath("//table[4]/tbody/tr[7]/td[2]/b/div").text
		sheet1.write(row, 9, answer)
		if check_exists_by_xpath("//table[4]/tbody/tr[5]/td[2]/div"):
			answer = answer + driver.find_element_by_xpath("//table[4]/tbody/tr[5]/td[2]/div").text
			sheet1.write(row, 9, answer)
			if check_exists_by_xpath("//table[4]/tbody/tr[3]/td[2]/div"):
				answer = answer + driver.find_element_by_xpath("//table[4]/tbody/tr[3]/td[2]/div").text
				sheet1.write(row, 9, answer)
		elif check_exists_by_xpath("//table[4]/tbody/tr[3]/td[2]/div"):
			answer = answer + driver.find_element_by_xpath("//table[4]/tbody/tr[3]/td[2]/div").text
			sheet1.write(row, 9, answer)
	elif check_exists_by_xpath("//table[4]/tbody/tr[5]/td[2]/div"):
		answer = driver.find_element_by_xpath("//table[4]/tbody/tr[5]/td[2]/div").text
		sheet1.write(row, 9, answer)
		if check_exists_by_xpath("//table[4]/tbody/tr[3]/td[2]/div"):
			answer = answer + driver.find_element_by_xpath("//table[4]/tbody/tr[3]/td[2]/div").text
			sheet1.write(row, 9, answer)
	elif check_exists_by_xpath("//table[4]/tbody/tr[3]/td[2]/div"):
		answer = driver.find_element_by_xpath("//table[4]/tbody/tr[3]/td[2]/div").text
		sheet1.write(row, 9, answer)


#Checks the 'Official' RFI box.  On this project, all RFIs are official, so there are no logic checks.		
def is_official(row):
	sheet1.write(row, 10, "1")	#The number 1 makes the box in prolog get checked off.
	
#checks to see if submittal has been closed by looking at date returned.  
#If it says (open) then it doesn't check the box.  Otherwise, it does.
def is_closed(row):
	if check_exists_by_xpath("//table[1]/tbody/tr[8]/td[2]"):
		date_ret = driver.find_element_by_xpath("//table[1]/tbody/tr[8]/td[2]").text
		open = "(open)"
		if date_ret == open:
			print("Status:          Open")
		else: 
			sheet1.write(row, 11, "1")
			print("Current Status:  Closed")
			
#Checks to see who reviewed the RFI.  Puts Kyle Lind if can't figure out.
def answered_by(row):
	if check_exists_by_xpath("//table[4]/tbody/tr[6]/td[2]"):
		reviewer = driver.find_element_by_xpath("//table[4]/tbody/tr[6]/td[2]").text
		if reviewer == "Reviewer - Brenda Powell":
			reviewer = reviewer.replace("Reviewer - ","")
			sheet1.write(row, 7, reviewer)		#Writes it to Excel Sheet
			sheet1.write(row, 6, "TLCTILD0")
		elif reviewer == "Reviewer - Kevin Casey":
			reviewer = reviewer.replace("Reviewer - ","")
			sheet1.write(row, 7, reviewer)		#Writes it to Excel Sheet
			sheet1.write(row, 6, "PA882249")
		else:
			reviewer = "Kyle Lind"
			sheet1.write(row, 7, reviewer)		#Writes it to Excel Sheet
			sheet1.write(row, 6, "HKSINC 3")
	elif check_exists_by_xpath("//table[4]/tbody/tr[4]/td[2]"):
		reviewer = driver.find_element_by_xpath("//table[4]/tbody/tr[4]/td[2]").text
		if reviewer == "Reviewer - Brenda Powell":
			reviewer = reviewer.replace("Reviewer - ","")
			sheet1.write(row, 7, reviewer)		#Writes it to Excel Sheet
			sheet1.write(row, 6, "TLCTILD0")
		elif reviewer == "Reviewer - Kevin Casey":
			reviewer = reviewer.replace("Reviewer - ","")
			sheet1.write(row, 7, reviewer)		#Writes it to Excel Sheet
			sheet1.write(row, 6, "PA882249")
		else:
			reviewer = "Kyle Lind"
			sheet1.write(row, 7, reviewer)		#Writes it to Excel Sheet
			sheet1.write(row, 6, "HKSINC 3")
	else:
		reviewer = "Kyle Lind"
		sheet1.write(row, 7, reviewer)		#Writes it to Excel Sheet
		sheet1.write(row, 6, "HKSINC 3")
	print("Reviewed by:     " + reviewer)	#Gives info to the user
			

	
#This loop iterates through each of the above functions for each RFI in the user specified range.
row = 1
for rfi in range(start_rfi, end_rfi+1):
	pullup_rfi(rfi, row)
	get_title(row)
	get_projAdmin(row)
	answered_by(row)
	get_dateSub(row)
	get_dateDue(row)
	get_dateRet(row)
	get_question(row)
	get_answer(row)
	is_official(row)
	is_closed(row)
	download_attachments()
	print("")
	row += 1

#converts start_rfi and end_rfi to strings so that we can use them to name the Excel file when we save it.
start_rfi = str(start_rfi)
end_rfi = str(end_rfi)
book.save('RFI ' + start_rfi + '-' + end_rfi + ' Women\'s Pavilion.xls')	#Save the Excel Sheet

raw_input("Press any key to exit.")
sys.exit(-1)
