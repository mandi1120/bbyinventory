##########    README    ##########
# This script mimics the VBA script format as of 12/11/17, which creates a new BBY Inventory file for each market
# by updating the market field in the Market tab causing the formulas to recalculate, then writes the data to a new file and saves.
# On 12/11/17 this script was timed at 59 minutes on my Dell computer for 87 markets. - MH 12/11/17
#
# Future additions:
# -Check if the file inadvertently closed during the for loop, and if so, re-open it and start at the most recent market
# -Calculate inventory numbers within python instead of using formulas to see if this speeds processing time.
##################################

import os, zipfile, re, shutil, xlwings, stat, openpyxl, time
from stat import S_IWUSR, S_IREAD
from openpyxl import load_workbook
import win32com.client

homedir = 'ahanway'
filename = 'C:\\Users\\'+ homedir + '\\Desktop\\BBY TV inventory - FSM email V2.xlsm'
#app = xlwings.apps[0]
wb = xlwings.Book(filename)
wb.app.screen_updating = False
#excel = win32com.client.Dispatch("Excel.Application")
start = time.time()
print("Process started at: " + time.asctime( time.localtime(time.time()) ) )

#inv file variables
adminWs = wb.sheets['Admin']
marketWs = wb.sheets['Market']
templateWs = wb.sheets['Template']
markets = adminWs.range('B16:B102').value

#temp workbook variables
temp_wb = xlwings.Book()
tempWbSh = temp_wb.sheets['Sheet1']

errors = []
thisMarket = 0
totalMarkets = 0
for eachMarket in markets:
	totalMarkets +=1

with open(filename):
	try:		
		for eachMarket in markets:
			try:
				thisMarket +=1
				
				if thisMarket <= totalMarkets:
				
					print(" - Creating " + eachMarket + " file (" + str(thisMarket) + "/" + str(totalMarkets) + ") ...")
					
					#copy template sheet to new workbook
					new_wb = xlwings.Book()
					newWbSh = new_wb.sheets[0]
					wb.sheets[templateWs].api.Copy(Before=newWbSh.api)
					new_wb.sheets[0].name = 'TV Inventory'
					new_wb.sheets['Sheet1'].delete()
			
					#update market field to re-calculate inventory
					marketWs.range('B4').value = eachMarket
			
					#paste values to temp workbook and sort by region/model/sku
					marketValues = marketWs.range('A3:AP100').value 
					tempWbSh.range('A1:AP98').value = marketValues		
					tempWbSh.range('A4:AP98').api.Sort(Key1=tempWbSh.range('F4').api, Order1=2,
						Key2=tempWbSh.range('C4').api, Order2=2,
						Key3=tempWbSh.range('B4').api, Order3=2, Orientation=1)
			
					#delete Sku column contents (column B)
					tempWbSh.range('B4:B100').clear()
			
					#cut/paste values from temp_wb to new_wb
					marketValuesTemp = tempWbSh.range('A1:AP100').value 
					new_wb.sheets['TV Inventory'].range('A1:AP100').value = marketValuesTemp
			
					#clear temp_wb data
					tempWbSh.clear()	

					#save and close market workbook
					new_wb.save(r'C:\\Users\\'+ homedir + '\\Documents\\BBY TV Inventory - FSM Report\\' + str(eachMarket) + ' BBY Inv.xlsx')
					new_wb.close()
					
			except Exception:
				#app = xlwings.App()
				open(filename,'r')
				wb = xlwings.Book(filename)
				temp_wb = xlwings.Book()
				tempWbSh = temp_wb.sheets['Sheet1']
				adminWs = wb.sheets['Admin']
				marketWs = wb.sheets['Market']
				templateWs = wb.sheets['Template']
				markets = adminWs.range('B16:B102').value
				#eachMarket = eachMarket-1
				errors = errors.append(eachMarket)
				print(" ----- ERROR: reopening file to continue")
				continue
				
	except Exception:
		print("Error")
	
wb.app.screen_updating = True
temp_wb.close()
end = time.time()
processingTime = round(((end - start)/60))
print("\n\nCompleted in " + str(processingTime) + " minutes.")
print( "  *** Files with errors: " + str(errors) )