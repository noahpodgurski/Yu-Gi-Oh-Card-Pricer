"""
Noah Podgurski 2019
"""
import tkinter
from tkinter import filedialog
import json
import requests #for grabbing HTML
import os
import xlrd #for reading excel file
import re
from bs4 import BeautifulSoup #for parsing HTML
import xlsxwriter #for writing to excel
import time
import warnings
warnings.filterwarnings("ignore") #ignore warnings

global total
global sellingTotal
global notFound
global myConditions
# start=time.clock()
myConditions = []

# print("""
# ██╗   ██╗██╗   ██╗       ██████╗ ██╗       ██████╗ ██╗  ██╗██╗                 
# ╚██╗ ██╔╝██║   ██║      ██╔════╝ ██║      ██╔═══██╗██║  ██║██║                 
#  ╚████╔╝ ██║   ██║█████╗██║  ███╗██║█████╗██║   ██║███████║██║                 
#   ╚██╔╝  ██║   ██║╚════╝██║   ██║██║╚════╝██║   ██║██╔══██║╚═╝                 
#    ██║   ╚██████╔╝      ╚██████╔╝██║      ╚██████╔╝██║  ██║██╗                 
#    ╚═╝    ╚═════╝        ╚═════╝ ╚═╝       ╚═════╝ ╚═╝  ╚═╝╚═╝                 
                                                                               
#  ██████╗ █████╗ ██████╗ ██████╗     ██████╗ ██████╗ ██╗ ██████╗███████╗██████╗ 
# ██╔════╝██╔══██╗██╔══██╗██╔══██╗    ██╔══██╗██╔══██╗██║██╔════╝██╔════╝██╔══██╗
# ██║     ███████║██████╔╝██║  ██║    ██████╔╝██████╔╝██║██║     █████╗  ██████╔╝
# ██║     ██╔══██║██╔══██╗██║  ██║    ██╔═══╝ ██╔══██╗██║██║     ██╔══╝  ██╔══██╗
# ╚██████╗██║  ██║██║  ██║██████╔╝    ██║     ██║  ██║██║╚██████╗███████╗██║  ██║
#  ╚═════╝╚═╝  ╚═╝╚═╝  ╚═╝╚═════╝     ╚═╝     ╚═╝  ╚═╝╚═╝ ╚═════╝╚══════╝╚═╝  ╚═╝  
# """)

def readFile(path):
	codes = []
	p = re.compile("([\\d]*[A-Z]+[\\d]*-[A-Z]*[\\d]*)")
	wb = xlrd.open_workbook(path)
	for sheet in wb.sheets():
	# sheet = wb.sheet_by_index(0)
		print("New Sheet")
		for x in range(0, sheet.nrows):
			for y in range(0, sheet.ncols):
				code = sheet.cell_value(x, y)
				if code.strip() != "":
					if (p.match(code)): #if its a code
						codes.append(code)
					# elif y == 1: #its a card condition
					# 	myConditions.append(code)
					else:
						print(f"COULDN'T FIND A CODE AT ({x}, {y})")
	codes.sort()
	print(codes)
	print(myConditions)
	print(f"{len(codes)} codes found")
	return codes

def scrapeData(codes):
	global total
	global myConditions
	global sellingTotal
	global notFound
	notFound = []
	data = []
	total = 0
	sellingTotal = 0
	for i in range(len(codes)):
		code = codes[i]
		if code == codes[i-1]:
			print("duplicate found")
			data.append(data[-1])
		else:
			req = requests.get(f'https://shop.tcgplayer.com/yugioh/product/show?advancedSearch=true&Number={code}&Price_Condition=Less+Than', verify=False).text
			print(f"Scraping from {code}")

			soup = BeautifulSoup(req, "html.parser")

			try:
				cardData = soup.find("div", {"class": "product__summary"}, "a").text
				cardData = cardData.split("\n")# looks like -> ['Blue-Eyes White Dragon', 'Starter Deck: Kaiba (YuGiOh)', 'Number SDK-001', 'Rarity Ultra Rare']
				cardData = [x for x in cardData if x != ""]
				cardData[2] = cardData[2][7:]
				cardData[3] = cardData[3][7:]

				listings = soup.findAll("div", {"class": "listing"})
				conditions = [listing.find("div", {"class": "condition"}).text for listing in listings]
				productPrice = [listing.find("span", {"class": "product-offer__price"}).text for listing in listings]

				cardData.append(conditions)
				cardData.append(productPrice)

				sellingTotal += float(productPrice[0][1:])
				price = soup.find("div", {"class": "product__prices"}).find("dd").text
				if type(price[1:]) != str:
					total += float(price[1:])
				cardData.append(price)
				data.append(cardData)
			except AttributeError:
				notFound.append(code)
				print(f"{code} not found")
				# print(myConditions)
				# print(len(myConditions))
				if len(myConditions) > 0 and myConditions[i]:
					myConditions.pop(i)


	# sort data by highest price
	def getKey(item):
		return item[-1]
	data.sort(key=getKey, reverse=True)
	return data

def worthMarket(condition, myCondition):
	if "damaged" in condition:
		condition = 0
	elif "heavily played" in condition:
		condition = 1
	elif "moderately played" in condition:
		condition = 2
	elif "lightly played" in condition:
		condition = 3
	elif "near mint" in condition or "like new" in condition:
		condition = 4
	elif "mint" in condition:
		condition = 5
	else:
		condition = 0

	if "damaged" in myCondition:
		myCondition = 0
	elif "heavily played" in myCondition:
		myCondition = 1
	elif "moderately played" in myCondition:
		myCondition = 2
	elif "lightly played" in myCondition:
		myCondition = 3
	elif "near mint" in myCondition or "like new" in myCondition:
		myCondition = 4
	elif "mint" in myCondition:
		myCondition = 5
	else:
		myCondition = 0

	if condition >= myCondition:
		return False
	return True

	


def writeData(data, path):
	# # sort data by highest price
	# def getKey(item):
	# 	return item[-1]
	# data.sort(key=getKey, reverse=True)

	fileLoc = os.path.dirname(path) + "/CardPricerOutput.xlsx"
	workbook = xlsxwriter.Workbook(fileLoc)
	worksheet = workbook.add_worksheet()
	worksheet.set_column('A:A', 63) #name
	worksheet.set_column('B:B', 24) #rarity
	worksheet.set_column('C:C', 30) #starter deck
	worksheet.set_column('D:D', 25) #card number
	worksheet.set_column('E:E', 18) #currently selling
	worksheet.set_column('F:F', 40) #selling condition
	worksheet.set_column('G:G', 25) #my condition
	worksheet.set_column('H:H', 17) #market price

	normalFormat = workbook.add_format({
		'font_name': 'Cambria',
		'font_size': 16,
		'border': 1,
		'align': 'center',
		'valign': 'vcenter',
		# 'fg_color': 'yellow',
		'text_wrap': 1,
	})
	headerFormat = workbook.add_format({
		'bold': True,
		'align': 'center',
		'font_size': 20,
		'font_name': 'Cambria',
		'border': 1,
		'text_wrap': 1
	})
	moneyFormat = workbook.add_format({
		'align': 'center',
		'font_name': 'Cambria',
		'font_size': 16,
		'num_format': '$#,##0.00',
		'top': 1,
		'left': 1,
		'right': 1,
		'bottom': 1,
		'bottom_color': '#000000',
		'top_color': '#000000',
		'left_color': '#000000',
		'right_color': '#000000'
	})
	boldMoneyFormat = workbook.add_format({
		'bold': True,
		'align': 'center',
		'font_name': 'Cambria',
		'font_size': 20,
		'num_format': '$#,##0.00',
		'top': 1,
		'left': 1,
		'right': 1,
		'bottom': 1,
		'bottom_color': '#000000',
		'top_color': '#000000',
		'left_color': '#000000',
		'right_color': '#000000'
	})

	worksheet.write(0, 0, "Name", headerFormat)
	worksheet.write(0, 1, "Rarity", headerFormat)
	worksheet.write(0, 2, "Starter Deck", headerFormat)
	worksheet.write(0, 3, "Card Number", headerFormat)
	worksheet.write(0, 4, "Currently Selling", headerFormat)
	worksheet.write(0, 5, "Selling Condition", headerFormat)
	worksheet.write(0, 6, "My Condition", headerFormat)
	worksheet.write(0, 7, "Market Price", headerFormat)

	if len(notFound) > 0:
		print(f"{len(notFound)} cards not found:")
		for x in notFound:
			print(x)
		print()

	for i in range(len(data)):
		print(f"Name: {data[i][0]} ({data[i][3]}), Market Price: {data[i][-1]} Currently selling {data[i][4][0]} at {data[i][5][0]}")
		worksheet.write(i+1, 0, data[i][0], normalFormat) #name
		worksheet.write(i+1, 1, data[i][3], normalFormat) #rarity
		worksheet.write(i+1, 2, data[i][1], normalFormat) #starter deck
		worksheet.write(i+1, 3, data[i][2], normalFormat) #card number
		worksheet.write(i+1, 5, data[i][4][0], normalFormat) #selling condition
		try:
			if len(myConditions) > 0 and myConditions[i]:
				worksheet.write(i+1, 6, myConditions[i], normalFormat) #my condition
				if worthMarket(data[i][4][0].lower(), myConditions[i].lower()):
					worksheet.write(i+1, 4, float(data[i][5][0][1:]), moneyFormat) #currently selling
					worksheet.write(i+1, 7, float(data[i][-1][1:]), boldMoneyFormat) #market price
				else:
					worksheet.write(i+1, 4, float(data[i][5][0][1:]), boldMoneyFormat) #currently selling
					worksheet.write(i+1, 7, float(data[i][-1][1:]), moneyFormat) #market price
			else:
				worksheet.write(i+1, 6, "", normalFormat) #my condition
				worksheet.write(i+1, 4, float(data[i][5][0][1:]), boldMoneyFormat) #currently selling
				worksheet.write(i+1, 7, float(data[i][-1][1:]), moneyFormat) #market price
		except:
			worksheet.write(i+1, 4, "UNAVAILABLE", normalFormat)
			worksheet.write(i+1, 7, "UNAVAILABLE", normalFormat)

	worksheet.write(len(data)+1, 4, sellingTotal, boldMoneyFormat) #currently selling total
	worksheet.write(len(data)+1, 7, total, boldMoneyFormat) #market price total
	print(f"Total: ${total}")


	workbook.close()
	print("Saved to {} \r\n".format(fileLoc))



def main():
	tkinter.Tk().withdraw() # Close the root window
	path = filedialog.askopenfilename(title = "Please select your excel file", filetypes = (("Excel files", "*.xlsx"), ("All files","*.*")))
	try:
		print(path)
	except:
		print("No Valid File selected, quitting.")
	writeData(scrapeData(readFile(path)), path)
main()


# print("Operation completed in {} seconds.".format(round((time.clock()-start), 2)))