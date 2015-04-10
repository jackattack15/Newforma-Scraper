from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import xlwt
import sys
import codecs
import string


if sys.stdout.encoding != 'cp850':
	sys.stdout = codecs.getwriter('cp850')(sys.stdout, 'strict')
if sys.stderr.encoding != 'cp850':
	sys.stderr = codecs.getwriter('cp850')(sys.stderr, 'strict')

#We need to know which RFIs to update, so ask the user
start_rfi = input('Please enter starting RFI #: ')
end_rfi = input('Please enter ending RFI #: ')

#Next log into Newforma
driver = webdriver.Chrome('C:\Python27\Lib\site-packages\selenium\webdriver\chrome\chromedriver.exe')
driver.get("https://projectcloud.newforma.com/")
login_email = driver.find_element_by_name('email')
login_email.send_keys("gsavage@brasfieldgorrie.com")
login_pswd = driver.find_element_by_name('pass')
login_pswd.send_keys("g021114")
login_pswd.send_keys(Keys.ENTER)
time.sleep(0.2)
driver.get("https://projectcloud.newforma.com/rfi/dashboard.php")
rfi_log = open('RFI' + str(start_rfi) + str(end_rfi), 'w')

#initiate Excel Workbook
book = xlwt.Workbook()
sheet1 = book.add_sheet('Sheet 1', cell_overwrite_ok=True)

style = xlwt.easyxf('font: bold 1')
sheet1.write(0, 0, 'Project RFI Number', style)
sheet1.write(0, 1, 'Subject', style)
sheet1.write(0, 2, 'Date Created', style)
sheet1.write(0, 3, 'Date Required', style)
sheet1.write(0, 4, 'Date Answered', style)
sheet1.write(0, 5, 'Authored By', style)
sheet1.write(0, 6, 'Question', style)
sheet1.write(0, 7, 'Answer', style)
#sheet1.write(0, 8, 'Answer 2', style)

def check_exists_by_xpath(xpath):
	try:
		driver.find_element_by_xpath(xpath).text
	except NoSuchElementException:
		return False
	return True

def check_exists_by_link(link):
	try:
		driver.find_element_by_partial_link_text(link)
	except NoSuchElementException:
		return False
	return True

def download_attachments():
	if check_exists_by_link('Response'):
		download = driver.find_element_by_partial_link_text('Response')
		download.click()
		print("Downloaded Response")
	
row = 1

#We will now create a loop that will search for each RFI in the range and scrape the pertinent information
for rfi in range(start_rfi, end_rfi+1):
	search = driver.find_element_by_name('search_text')
	search.send_keys(rfi)
	search.send_keys(Keys.RETURN)
	search_rfi = str(rfi)
	print search_rfi

	if check_exists_by_link('RFI-0' + search_rfi):
		current_rfi = driver.find_element_by_partial_link_text('RFI-0' + search_rfi)
		current_rfi.click()
	
	
	#Insert Function to scrape webpage
	sheet1.write(row, 0, search_rfi)
	
	if check_exists_by_xpath("//table[1]/tbody/tr[3]/td[2]"):
		title = driver.find_element_by_xpath("//table[1]/tbody/tr[3]/td[2]").text
		sheet1.write(row, 1, title)
	
	if check_exists_by_xpath("//table[1]/tbody/tr[5]/td[2]"):
		due_date = driver.find_element_by_xpath("//table[1]/tbody/tr[5]/td[2]").text
		sheet1.write(row, 3, due_date)
	
	if check_exists_by_xpath("//table[1]/tbody/tr[6]/td[2]"):
		date_sub = driver.find_element_by_xpath("//table[1]/tbody/tr[6]/td[2]").text
		sheet1.write(row, 2, date_sub)
	
	if check_exists_by_xpath("//table[1]/tbody/tr[7]/td[2]"):
		date_ret = driver.find_element_by_xpath("//table[1]/tbody/tr[7]/td[2]").text
		sheet1.write(row, 4, date_ret)
	
	if check_exists_by_xpath("//table[1]/tbody/tr[9]/td[2]"):
		question = driver.find_element_by_xpath("//table[1]/tbody/tr[9]/td[2]").text
		sheet1.write(row, 6, question)
	
	if check_exists_by_xpath("//table[4]/tbody/tr[5]/td[2]/b/div"):
		answer =  driver.find_element_by_xpath("//table[4]/tbody/tr[5]/td[2]/b/div").text
		sheet1.write(row, 7, answer)
	elif check_exists_by_xpath("//table[4]/tbody/tr[7]/td[2]/div"):
		answer =  driver.find_element_by_xpath("//table[4]/tbody/tr[7]/td[2]/div").text
		sheet1.write(row, 7, answer)
	if check_exists_by_xpath("//table[4]/tbody/tr[9]/td[2]/div"):
		answer = answer + driver.find_element_by_xpath("//table[4]/tbody/tr[9]/td[2]/div").text
		sheet1.write(row, 7, answer)
	
	if check_exists_by_xpath("//table[3]/tbody/tr[4]/td[2]"):
		proj_admin = driver.find_element_by_xpath("//table[3]/tbody/tr[4]/td[2]").text
		proj_admin = proj_admin.replace("Project Admin (Contractor) - ","")
		sheet1.write(row, 5, proj_admin)
		
	download_attachments()
	
	
	print(title)
	print("Project Admin:   " + proj_admin)
	print("Date Submitted:  " + date_sub)
	print("Date Due:        " + due_date)
	print("Date Returned:   " + date_ret)
	print("")
#	print(question)
	
	row += 1
	
	
	
	#Insert Function to input info into prolog

book.save('Auto RFI.xls')

