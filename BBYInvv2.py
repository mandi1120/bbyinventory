##########    README    ##########
# This script filters the full TV inventory down to each store in each market, then writes the market inventory to a new file. 
# This script replaces the previously used macro which updated the market field in the BBY Inventory file, causing the 
# sumifs formulas to recalculate, then wrote to a new file, which was very slow to process.  
# MH - 1/3/2018
##################################

import os, xlwings, time, xlrd, openpyxl, win32com.client
import pandas as pd
import numpy as np
from pandas import Series, DataFrame
from openpyxl import Workbook, load_workbook

excel = win32com.client.Dispatch("Excel.Application")
excel.ScreenUpdating = False

### set your directory here
homedir = 'ahanway'

### set the path to the BBY Inventory file here; saving first in case you forgot to
filename = 'C:\\Users\\'+ homedir + '\\Desktop\\BBY TV inventory - FSM email V2-3.xlsm'
wb = xlwings.Book(filename)
wb.save()

#using an excel macro to unhide all sheets in the bby inv file
#if the sheets remained hidden, an error is generated in the final market file
a = wb.api.Application.Run("unhideSheets")

#starting script and timer
start = time.time()
print("\n\n    ---------------->   Process started: " + time.asctime( time.localtime(time.time()) ) + "   <---------------- \n")

#the original bby inv file is too large and slows the script processing time; create a new temp saved to desktop
#with only admin, bby inv email, list, and template tabs from original bby inv file to pull from in the script.
print("   - Step 1: Creating valuedWb temp file saved as: \Desktop\BBY_TV_Inv_Temp.xlsx...")
print("         - This temp file will be deleted from the desktop upon script completion. This may take a few minutes...")
valuedWb = xlwings.Book()
wb.sheets['Template'].api.Copy(Before=valuedWb.sheets[0].api)
valuedWb.sheets.add()
valuedWb.sheets.add()
valuedWb.sheets["Sheet1"].range("A1").value = wb.sheets['Admin'].range('A1:G200').value
valuedWb.sheets["Sheet1"].name = 'Admin'
valuedWb.sheets["Sheet3"].range("A1").value = wb.sheets['List'].range('A1:N1000').value
valuedWb.sheets["Sheet3"].name = 'List'
valuedWb.sheets["Sheet4"].range("A1").value = wb.sheets['BBY Inv email'].range('A1:K40000').value
valuedWb.sheets["Sheet4"].name = 'BBY Inv email'
valuedWb.save('C:\\Users\\'+ homedir + '\\Desktop\\BBY_TV_Inv_Temp.xlsx')
valuedWbFilename = 'C:\\Users\\'+ homedir + '\\Desktop\\BBY_TV_Inv_Temp.xlsx'
			
### creating connection to another temp workbook used to manipulate market data before later writing it to the final file
#tempwb = xlwings.Book('C:\\Users\\'+ homedir + '\\Desktop\\temp.xlsx') - saves temp file, used for testing
tempwb = xlwings.Book()
tempWbSh = tempwb.sheets['Sheet1']

#turn bby inv email sheet into a data frame, keep only this year & last year skus: KU,KS,LS,MU,QN,WMN
#temp bby inv file must be closed before reading into data frame; also closing main wb so it doesn't slow processing
print("   - Step 2: Creating BBY inv email data frame...")
wb.close()
valuedWb.close()
invxlfile = pd.ExcelFile(valuedWbFilename)
df = pd.read_excel(invxlfile, 'BBY Inv email')
df = df[['Region','Market','StoreId','Warehouse?','Sku','Inventory']]
df = df [ df['Sku'].str.contains("KU|KS|LS|MU|WMN|QN") ]

#create data frame from store list
print("   - Step 3: Creating store list data frame...")
storeListDf = pd.read_excel(invxlfile, 'List')
storeListDf = storeListDf[['StoreId','City','Tier','Type','DC','DDC']]
storeListDf = storeListDf.rename(columns={'DC': 'DC Store#', 'DDC':'DDC Store#'})

#create national warehouse DC/DDC inventory totals in new data frame
print("   - Step 4: Finding national DC/DDC inventory numbers...")
whseInvColumn = df [ df['Warehouse?'] == "Whse"]
whseInvColumn = whseInvColumn.groupby(by=['Sku'])['Inventory'].sum()
whseInvColumn = pd.DataFrame({'Inventory':whseInvColumn.values,'Sku':whseInvColumn.index})
whseInvColumnNewIndex = whseInvColumn.set_index('Sku')

#add National DC/DDC column to df using Sku as index
merged = df.merge(whseInvColumn, how='left', left_on='Sku', right_on='Sku')
merged = merged.rename(columns={'Inventory_y': 'National DC/DDC Inventory', 'Inventory_x':'Units'})
merged = merged.merge(storeListDf, how='left', left_on='StoreId', right_on='StoreId')

#create list of all unique skus to be used to add in any missing (0 inv) skus to the market file
fullSkuList = merged [[ 'Sku', 'National DC/DDC Inventory' ]]
fullSkuList = fullSkuList.drop_duplicates('Sku')
fullSkuList = fullSkuList.values.tolist()

#concatenate StoreId with City, Tier, and Type into a new column to use in the final pivot table
merged["StoreIdMerged"] = merged["StoreId"] + str("\\") + merged["City"] + str("\\Tier-") + merged["Tier"].map(lambda x: "{:.0f}".format(x)) + str(" ") + merged["Type"]	

#reopen the temp bby inv file and set variables for sheets
valuedWb = xlwings.Book(valuedWbFilename)
adminWs = valuedWb.sheets['Admin']
templateWs = valuedWb.sheets['Template']

#setting variable based on "units as of" field in Admin sheet
unitsAsOfDate = adminWs.range('A11:A11').value
	
#set range for market list in admin tab
markets = adminWs.range('B23:B109').value
thisMarket = 0
totalMarkets = 0
for eachMarket in markets:
	totalMarkets +=1
	
#begin to iterate through each market
print("   - Step 5: Beginning to loop through markets...")
for eachMarket in markets:
	thisMarket +=1
	
	if thisMarket <= totalMarkets:
	
		print("         -  Creating " + eachMarket + " file (" + str(thisMarket) + "/" + str(totalMarkets) + ") ...")
		
		#create a new df of only this market stores' inventory				
		marketdf = merged[ merged['Market'] == eachMarket]
	
		#merge full sku list into the market df listing "0" if no inventory in this market
		#turn the market skus into list
		#check if each is in the full sku list; if not, append it to the missing sku list
		marketSkuList = marketdf['Sku'].tolist()
		missingSkuList = []
		for eachSkuList in fullSkuList:
			if eachSkuList[0] not in marketSkuList:
				missingSkuList.append(eachSkuList)
		#turn missingskulist into dataframe
		missingSkuDf = pd.DataFrame(missingSkuList)
		missingSkuDf = missingSkuDf.fillna(0)
		missingSkuDf.columns = ['Sku', 'National DC/DDC Inventory']

				
		#create Market inventory totals per sku
		marketInvColumn = marketdf.groupby(by=['Sku'])['Units'].sum()
		marketInvColumn = pd.DataFrame({'Units':marketInvColumn.values,'Sku':marketInvColumn.index})

		#add market totals to market df
		merged2 = marketdf.merge(marketInvColumn, how='left', left_on='Sku', right_on='Sku')
		merged2 = merged2.rename(columns={'Units_y': (str(eachMarket) + '\\Covered\\Stores\\Units'), 'Units_x':'Units'})

		#add 2 new columns to merged2 df for sku model & size
		models = []
		sizes = []
		for row in merged2['Sku']:
			if row[:3] == "WMN":
				models.append(row[:7])
				sizes.append("N/A")
			else:
				models.append(row[4:10])
				sizes.append(str(row[2:4] + '"'))				
		merged2['Model'] = models
		merged2['Size'] = sizes			

		#add 2 new columns to missingSkuDf for sku model & size
		models1 = []
		sizes1 = []
		for row in missingSkuDf['Sku']:
			if row[:3] == "WMN":
				models1.append(row[:7])
				sizes1.append("N/A")
			else:
				models1.append(row[4:10])
				sizes1.append(str(row[2:4] + '"'))				
		missingSkuDf['Model'] = models1
		missingSkuDf['Size'] = sizes1	
		missingSkuDf = missingSkuDf [['Sku','Model','Size','National DC/DDC Inventory']]
		missingSkuDf.sort_values(['Model', 'Size'], ascending=[False, False])

		print(missingSkuDf)
		
		#sum DC/DDC inventory for each sku based on what DC/DDC's match up to each store
		merged3 = merged2

		for row in merged2:
			for row in merged2['DC Store#']:
				x = len(row)
				dc = row[0:x]
			for row in merged2['DDC Store#']:
				y = len(row)
				ddc = row[0:y]

			storeDcDDCInv = merged2[(dc == merged.StoreId) | (ddc == merged.StoreId)]
			storeDcDDCInv = storeDcDDCInv.groupby(by=['Sku'])['Units'].sum()
			storeDcDDCInv = pd.DataFrame({'Units':storeDcDDCInv.values,'Sku':storeDcDDCInv.index})
			
			#rename this data frame's Units as DC/DDC Units - starting with "z" to pull into pivot df in correct order after aUnits
			storeDcDDCInv = storeDcDDCInv.rename(columns={'Units': 'zDC/DDC Units'}) 
		
			#add column for storeDcDDCInv
			merged4 = merged3.merge(storeDcDDCInv, how='left', left_on='Sku', right_on='Sku')
			
		#rename Units - starting with "a" to pull into pivotdf in correct order before zDC/DDC Units	
		merged4 = merged4.rename(columns={'Units':'aUnits'}) 
		
		#turn market inventory df into a pivot to get in correct market report format
		pivotdf = merged4.pivot_table(index=['Region','Market','Sku', 'Model', 'Size','National DC/DDC Inventory',
			(str(eachMarket) + '\\Covered\\Stores\\Units')], columns='StoreIdMerged',values=['zDC/DDC Units','aUnits'], fill_value=0).reset_index()
			
		#sort df while headings are still in first row by market inventory (descending)
		pivotdf = pivotdf.sort_values(by=[(str(eachMarket) + '\\Covered\\Stores\\Units'), 'Model', 'Sku'], ascending=[0,0,0])
		
		#reorder column labels moving storeId to row 1
		pivotdf = pivotdf.reorder_levels([1, 0], axis=1)
		pivotdf = pivotdf.sort_index(level=[1],axis=1, ascending=[False], na_position='first')

		#rename Units & DC/DDC Units columns to remove first character that was used for sorting
		pivotdf = pivotdf.rename(columns={'zDC/DDC Units': 'DCs', 'aUnits': 'Units'})
		
		
		####### TEMP FILE SECTION #######
		#write pivot df to temp excel doc
		tempWbSh.range('A1').value = pivotdf
		
		#clear contents of column A
		tempWbSh.range("A:A").api.EntireColumn.Clear()

		#Set variables for region and market, then delete columns
		market = tempWbSh.range("C5").value
		region = tempWbSh.range("B5").value
		tempWbSh.range("B:B,C:C").api.EntireColumn.Delete()
		
		#insert two new rows under first row
		tempWbSh.range("2:3").api.EntireRow.Insert()	

		#split field into 4 rows: National | DC/DDC | (blank) | Inventory (E1:E4)	
		natDcDDCLabel = tempWbSh.range("E4").value
		natLabel, dcDDCLabel, invLabel = natDcDDCLabel.split()
		tempWbSh.range("E1").value = natLabel
		tempWbSh.range("E2").value = dcDDCLabel
		tempWbSh.range("E4").value = invLabel
		
		#split field into 4 rows: {Market} | Covered | Store | Units (F1:F4)
		marketInvLabel = tempWbSh.range("F4").value
		marketLabel, coveredLabel, storeLabel, unitsLabel = marketInvLabel.split('\\')
		tempWbSh.range("F1").value = marketLabel
		tempWbSh.range("F2").value = coveredLabel
		tempWbSh.range("F3").value = storeLabel
		tempWbSh.range("F4").value = unitsLabel
		
		#split StoreIdMerged field into 3 rows: ID | City | Tier&Type
		storeLabelRange = tempWbSh.range("G1:AP1")		
		for store in storeLabelRange:
			storeValue = tempWbSh.range(store).value
			if storeValue is not None:
				storeColumn = store.column
				storeIdMergedLabel = tempWbSh.range(store).value
				storeIdLabel, cityLabel, tierTypeLabel = storeIdMergedLabel.split('\\')
				tempWbSh.range(1, storeColumn).value = storeIdLabel
				tempWbSh.range(2, storeColumn).value = cityLabel
				tempWbSh.range(3, storeColumn).value = tierTypeLabel

		#Add "Region" to cell A1-C1, "Market" to cell A2-C2, "Units as of:" to Cell A3-A4
		tempWbSh.range("A1").value = "Region:"
		tempWbSh.range("C1").value = region
		tempWbSh.range("A2").value = "Market:"
		tempWbSh.range("C2").value = market
		tempWbSh.range("A3").value = "Units as of:"
		tempWbSh.range("A4").value = unitsAsOfDate
		tempWbSh.range("C3").value = "Products"
		
		#create new row with only store#, create list from store column to sort stores left-to-right
		tempWbSh.range("1:1").api.EntireRow.Insert()
		storeIdRange = tempWbSh.range("G2:AP2")	
		storeList = []			
		for store in storeIdRange:
			storeIdValue = tempWbSh.range(store).value
			if storeIdValue is not None:
				storeIdColumn = store.column
				storeIdNumber = tempWbSh.range(store).value
				tempWbSh.range(1, storeIdColumn).value = storeIdNumber[4:len(storeIdNumber)]
				storeList.append(tempWbSh.range((1, storeIdColumn),(101,storeIdColumn)).value)
		
		#sort columns by Units/DC's label (descending), then by storeId (ascending)
		storeList = sorted(sorted(storeList, key = lambda x : x[4], reverse=True), key = lambda x: (x[0]))
		
		#write the sorted store columns (G-AP) to the file
		for n in range(0,len(storeList)):
			r4 = tempWbSh.range((1, n+7), (101, n+7))
			r4.options(transpose = True).value = storeList[n]
			
		#where market inventory > 0, sort skus by series / size
		tempWbSh.range('B6:AP101').api.Sort(Key1=tempWbSh.range('C6').api, Order1=2, Key2=tempWbSh.range('D6').api, Order2=1, Orientation=1)
		
		#add missingSkuList into the file then sort by series / size
		#fill in values for each store as "0"
		#find last sku row, the paste missingskudf under it with natdcddcinv, model, size columns
		missingSkus = missingSkuDf.values
		missingSkuRng = tempWbSh.range("B6:B101")
		for cell in missingSkuRng:
			colValue = tempWbSh.range(cell).value
			if colValue is None:
				thisColumn = cell.column
				thisRow = cell.row
				tempWbSh.range(thisRow, thisColumn).value = missingSkus
				break
		
		#fill in blank cells with 0
#		checkEmptiesRng = tempWbSh.range("F2:AP2")	
#		skuRange2 = tempWbSh.range('B5:B101')
#		for each in checkEmptiesRng:
#			checkEmptiesValue = tempWbSh.range(each).value
#			if checkEmptiesValue is not None:
#				thisColumn = each.column
#				
#				for cell in skuRange2:
#					skuValue = tempWbSh.range(cell).value
#					if skuValue is not None:
#						skuRow = cell.row
#						unitsRange = tempWbSh.range(skuRow,thisColumn)
#						
#						#for unitsCell in unitsRange:
#						unitsValue = tempWbSh.range(unitsRange).value
#						if unitsValue is None:
#							unitsValue = "0"
				
		#sort market skus by 

		#where market inventory = 0, sort skus by series / size

		
		###### FINAL MARKET WORKBOOK SECTION ######
		#create a new workbook from the template tab, rename sheet as "TV Inventory", delete unused Sheet1
		new_wb = xlwings.Book()
		newWbSh = new_wb.sheets[0]
		valuedWb.sheets[templateWs].api.Copy(Before=newWbSh.api)
		new_wb.sheets[0].name = 'TV Inventory'
		new_wb.sheets['Sheet1'].delete()
		
		#copy data from temp file to new workbook, then set freeze panes
		marketValuesTemp = tempWbSh.range('A2:AP101').value 
		new_wb.sheets['TV Inventory'].range('A1:AP100').value = marketValuesTemp
		new_wb.FreezePanes = False
		new_wb.sheets['TV Inventory'].range("G5").select()
		new_wb.FreezePanes = True
		
		#save new workbook as "[market] BBY TV Inventory.xlsm" to folder: C:\Users\username\Documents\BBY TV Inventory - FSM Report\
		new_wb.save(r'C:\\Users\\'+ homedir + '\\Documents\\BBY TV Inventory - FSM Report\\' + str(eachMarket) + ' BBY Inv.xlsx')
		new_wb.close()
		
		#clear data from temp workbook; keep it open for the next iteration
		tempWbSh.clear()	

	
#close/delete temp files 
tempwb.close()
valuedWb.close()
os.remove(valuedWbFilename)

#end the timer and report processing speed
end = time.time()
processingTime = round(((end - start)/60))
print("\n\n    ---------------->   Completed in " + str(processingTime) + " minutes. \n")
print("    ---------------->   Now use the Excel macro to email files to FSMs. \n")
print("    ---------------->   Reopening the BBY Inventory file. \n\n")

#reopening the bby inv file so you can run the email macro when you're ready
excel.ScreenUpdating = True
os.startfile(filename)
