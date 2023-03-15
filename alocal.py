from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
from openpyxl import load_workbook
import time, sys, re




def startUp(driver):
	driver.get("https://alocaldevelopment.com/")


def logIn(username, password, driver):
	user = driver.find_element(By.XPATH, "//input[1]")
	user.send_keys(username)
	passw = driver.find_element(By.XPATH, "//input[@name='password']")
	passw.send_keys(password)
	passw.send_keys(Keys.RETURN)

def searchBy(driver):
	inputs = getInputs(driver)
	inputs[1].send_keys(Keys.RETURN)
	inputs[1].send_keys(Keys.UP)
	inputs[1].send_keys(Keys.RETURN)


#enters the search criterion for first input on html page
def searchCriteria(driver):
	criteria = driver.find_element(By.XPATH, "//input[1]")
	criteria.send_keys("fgdfg")
	criteria.send_keys(Keys.RETURN)

#finds zipcode input and enters desired zipcode
def enterZipcode(zipcode, driver):
	zipelem = driver.find_element(By.XPATH, "//input[@name='zipCode']")
	zipelem.clear()
	zipelem.send_keys(zipcode)
	zipelem.send_keys(Keys.TAB)
	
def enterNAICSNumber(NAICSNumber, inputs, wait):
	inputs[4].clear()
	inputs[4].send_keys(NAICSNumber)
	wait.until(EC.presence_of_all_elements_located((By.XPATH, "//input[@aria-expanded='true']")))
	inputs[4].send_keys(Keys.DOWN)
	
	inputs[4].send_keys(Keys.RETURN)

def activateSearch(driver):
	search = driver.find_element(By.XPATH, "//button[1]")
	search.send_keys(Keys.RETURN)

def getInputs(driver):
	return driver.find_elements(By.XPATH, "//input[@role='combobox']")



def naicsSelect(naicsInput, inputs):
	inputs[naicsInput].send_keys(Keys.ENTER)

def naicsClear(naicsInput, inputs):
	inputs[naicsInput].clear()

def naicsDown(naicsInput, inputs):
	inputs[naicsInput].send_keys(Keys.DOWN)

def naicsDelete(naicsInput, inputs):
	inputs[naicsInput].send_keys(Keys.DELETE)

def naicsEscape(naicsInput, inputs):
	inputs[naicsInput].send_keys(Keys.ESCAPE)
	
	

def findCellsIndustries(driver):
	return driver.find_elements(By.XPATH, "//mat-cell[@class='mat-cell cdk-column-industry mat-column-industry ng-tns-c14-1 ng-star-inserted']")

def findCellsDemand(driver):
	return driver.find_elements(By.XPATH, "//mat-cell[@class='mat-cell cdk-column-surplusShortage mat-column-surplusShortage ng-tns-c14-1 ng-star-inserted']")




#Strips out marked option in NAICS combobox from the NAICS digit and leading spaces to be comparable to industry names in the industry table
def stripMarkedString(driver, naicsInput, inputs):
	try:	
		option = driver.find_elements(By.XPATH, "//ng-dropdown-panel/div/div/div[@class='ng-option ng-star-inserted ng-option-marked'] | //ng-dropdown-panel/div/div/div[@class='ng-option ng-option-marked ng-star-inserted']")
		return re.sub(r'[0-9][0-9]-[0-9][0-9]|[0-9]', '' ,option[0].text).lstrip()
	except:
		print("no element")
		naicsDown(naicsInput, inputs)
		option = driver.find_elements(By.XPATH, "//ng-dropdown-panel/div/div/div[@class='ng-option ng-star-inserted ng-option-marked'] | //ng-dropdown-panel/div/div/div[@class='ng-option ng-option-marked ng-star-inserted']")
		return re.sub(r'[0-9][0-9]-[0-9][0-9]|[0-9]', '' ,option[0].text).lstrip()



def naicsNumber(driver, naicsInput, inputs):
	try:
		option = driver.find_elements(By.XPATH, "//ng-dropdown-panel/div/div/div[@class='ng-option ng-star-inserted ng-option-marked'] | //ng-dropdown-panel/div/div/div[@class='ng-option ng-option-marked ng-star-inserted']")
		match = re.search(r'[0-9]+', option[0].text)
		return match.group(0)
	except:
		naicsDown(naicsInput, inputs)		
		option = driver.find_elements(By.XPATH, "//ng-dropdown-panel/div/div/div[@class='ng-option ng-star-inserted ng-option-marked'] | //ng-dropdown-panel/div/div/div[@class='ng-option ng-option-marked ng-star-inserted']")
		match = re.search(r'[0-9]+', option[0].text)
		return match.group(0)



#Used to store the name and net demand of industries at 2 digit NAICS that have a net demand greater than 0 
def makeDictionary(driver):
	industryElems = findCellsIndustries(driver)
	demandElems = findCellsDemand(driver)
	dictionary = {}
	
	for i in range(0, len(industryElems)):
		#if(float(demandElems[i].text) >= netdemand):
		dictionary[industryElems[i].text]= float(demandElems[i].text)
	return dictionary


def selectOption(inputNum, inputs, driver, wait):
	elements = findCellsIndustries(driver)
	naicsSelect(inputNum, inputs)
	wait.until(EC.presence_of_all_elements_located((By.XPATH, "//div[@role='option']")))
	naicsDown(inputNum, inputs)
	naicsSelect(inputNum, inputs)
	activateSearch(driver)
	try:
		wait.until(EC.staleness_of(elements[0]))
	except:
		time.sleep(2)

