'''     
  README:
  * This script replaces the Excel macro previously used in the BBY TV Inventory Report file, which created separate 
  workbooks listing sku inventory per store in each market, by updating the Market field in the Market tab causing the 
  formulas to recalculate and then writing the data to a new file, which was very slow to process. 
  * This version filters inventory by market and then calculates market total inventory and DC/DDC totals per sku per store 
  in pandas dataframes, writes the market data as a pivot table to a new file, adds report formatting, and then saves 
  the file in the folder "\Documents\BBY TV Inventory Market Reports". 
  * Email handling is currently still done using the Excel macro emailReports, which attaches each market file 
  from the BBY TV Inventory Market Reports folder to a message and emails it to the corresponding FSM.
  * Tested as 10 minutes to complete 87 market files.
  MH - 1/3/2018
  
'''

import os, xlwings, time, xlrd, win32com.client
import pandas as pd
import numpy as np
from pandas import Series, DataFrame
from operator import itemgetter

#excel = win32com.client.Dispatch("Excel.Application")
#excel.ScreenUpdating = False

### set your directory here
homedir = 'ahanway'

### set the path to the BBY Inventory file here; saving first in case you forgot to
filename = 'C:\\Users\\'+ homedir + '\\Desktop\\BBY TV inventory - FSM email V2.xlsm'
wb = xlwings.Book(filename)
wb.save()

#using an excel macro to unhide all sheets in the bby inv file
#because if the sheets remained hidden, an error is generated in the final market file
a = wb.api.Application.Run("unhideSheets")

#starting script and timer
start = time.time()
print("\n\n    ---------------->   Process started: " + time.asctime( time.localtime(time.time()) ) + "   <---------------- \n")

#the original bby inv file is too large and slows the script processing time; create a new temp saved to desktop
#with only admin, bby inv email, list, and template tabs from original bby inv file to pull from in the script.
print("   - Step 1: Creating valuedWb temp file saved as: \Desktop\BBY_TV_Inv_Temp.xlsx...")
print("         - File will be deleted from the desktop upon script completion. This part may take a few minutes...")
valuedWb = xlwings.Book()
wb.sheets['Template'].api.Copy(Before=valuedWb.sheets[0].api)
valuedWb.sheets.add()
valuedWb.sheets.add()
valuedWb.sheets["Sheet1"].range("A1").value = wb.sheets['Admin'].range('A1:G200').value
valuedWb.sheets["Sheet1"].name = 'Admin'
valuedWb.sheets["Sheet3"].range("A1").value = wb.sheets['List'].range('A1:N1000').value
valuedWb.sheets["Sheet3"].name = 'List'
valuedWb.sheets["Sheet4"].range("A1").value = wb.sheets['BBY Inv email'].range('A1:K80000').value
valuedWb.sheets["Sheet4"].name = 'BBY Inv email'
valuedWb.save('C:\\Users\\'+ homedir + '\\Desktop\\BBY_TV_Inv_Temp.xlsx')
valuedWbFilename = 'C:\\Users\\'+ homedir + '\\Desktop\\BBY_TV_Inv_Temp.xlsx'
			
### creating connection to another temp workbook used to manipulate market data before later writing it to the final file
#tempwb = xlwings.Book('C:\\Users\\'+ homedir + '\\Desktop\\temp.xlsx') #- saves temp file, used for testing
tempwb = xlwings.Book()
tempWbSh = tempwb.sheets['Sheet1']


#turn bby inv email sheet into a data frame, keep only this year & last year skus: KU,KS,LS,MU,QN,WMN
#closing temp bby inv file because must be closed before reading into data frame; also closing main wb so it doesn't slow processing
print("   - Step 2: Creating BBY inv email dataframe...")
wb.close()
valuedWb.close()
invxlfile = pd.ExcelFile(valuedWbFilename)
df = pd.read_excel(invxlfile, 'BBY Inv email')
df = df[['Region','Market','StoreId','Warehouse?','Sku','Inventory']]
df = df [ df['Sku'].str.contains("LS|MU|NU|WMN|QN") ]


#create data frame from store list
print("   - Step 3: Creating store list dataframe...")
storeListDf = pd.read_excel(invxlfile, 'List')
storeListDf = storeListDf[['StoreId','City','Tier','Type','DC','DDC']]
storeListDf = storeListDf.rename(columns={'DC': 'DC Store#', 'DDC':'DDC Store#'})


#####  National DC/DDC Inventory Column Section  #####
#create national warehouse DC/DDC inventory totals in new data frame
print("   - Step 4: Finding national DC/DDC inventory numbers...")
whseInvDf = df [ df['Warehouse?'] == "Whse"]
whseInvColumn = whseInvDf.groupby(by=['Sku'])['Inventory'].sum()
whseInvColumn = pd.DataFrame({'Inventory':whseInvColumn.values,'Sku':whseInvColumn.index})
whseInvColumnNewIndex = whseInvColumn.set_index('Sku')
#add National DC/DDC column to df using Sku as index
merged = df.merge(whseInvColumn, how='left', left_on='Sku', right_on='Sku')
merged = merged.rename(columns={'Inventory_y': 'National DC/DDC Inventory', 'Inventory_x':'Units'})
merged = merged.merge(storeListDf, how='left', left_on='StoreId', right_on='StoreId')

#####   Sku & Size Column Section   #####
#add 2 new columns to merged df for sku model & size
models = []
sizes = []
for row in merged['Sku']:
	if row[:3] == "WMN":
		models.append(row[:7])
		sizes.append("N/A")
	else:
		models.append(row[4:10])
		sizes.append(str(row[2:4] + '"'))				
merged['Model'] = models
merged['Size'] = sizes			


#####  List of all unique skus  #####
#create list of all unique skus used to add any missing (0 inv) skus to the market file
fullSkuList = merged [[ 'Sku', 'Model', 'Size', 'National DC/DDC Inventory' ]]
fullSkuList = fullSkuList.drop_duplicates('Sku')
fullSkuList = fullSkuList.values.tolist()



#####  StoreIdMerged Field   #####
#concatenate StoreId with City, Tier, and Type into a new column to use in the final pivot table
merged["StoreIdMerged"] = merged["StoreId"] + str("\\") + merged["City"] + str("\\Tier-") + merged["Tier"].map(lambda x: "{:.0f}".format(x)) + str(" ") + merged["Type"]	


#####  BBY Inv File Variables #####
#reopen the temp bby inv file and set variables for sheets
valuedWb = xlwings.Book(valuedWbFilename)
adminWs = valuedWb.sheets['Admin']
templateWs = valuedWb.sheets['Template']

#setting variable based on "units as of" field in Admin sheet
unitsAsOfDate = adminWs.range('A12').value
	
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
		#(skus that do not have inventory in this market are not included)		
		marketdf = merged[ merged['Market'] == eachMarket]
	
		
		#####   Adding Missing Skus to Market DF   #####		
		#find all stores / skus in this market
		#create list of all stores in this market
		marketStores = marketdf.drop_duplicates('StoreId')
		marketStores = marketStores[['Region','Market','StoreId','Warehouse?','City','Tier','Type','DC Store#','DDC Store#','StoreIdMerged']].values
		marketStoresList = marketStores.tolist()
		#add each sku in fullskulist to the missing sku list (listofskustore) for each store
		listofskustore = []
		for eachMarketStore in marketStoresList:
			for eachSkuList in fullSkuList:
				#if eachSkuList[0] not in storeSkuList:
				missingskuAndStore = eachMarketStore + eachSkuList				
				listofskustore.append(missingskuAndStore)
				
		#convert listofskustore into df ---#this contains all skus with store info but no units
		missingSkuheaders = ['Region','Market','StoreId','Warehouse?','City','Tier','Type','DC Store#','DDC Store#','StoreIdMerged','Sku', 'Model', 'Size', 'National DC/DDC Inventory']
		missingSkuDf = pd.DataFrame(listofskustore, columns=missingSkuheaders)
		
		#since missingskudf has all skus, add in units column from market, then replace what's already in marketdf
		#this lists both 0 inventory skus with skus in market
		marketdf = marketdf [['StoreId','Sku','Units']]
		marketdf2 = missingSkuDf.merge(marketdf, how='left', left_on=['Sku','StoreId'], right_on=['Sku','StoreId'])

		
		
		#######   Market's Total Inventory For All Stores Column   #######
		#create Market's inventory totals for all stores per sku
		marketInvColumn = marketdf2.groupby(by=['Sku'])['Units'].sum()
		marketInvColumn = pd.DataFrame({'Units':marketInvColumn.values,'Sku':marketInvColumn.index})

		#add market totals column to market df
		merged2 = marketdf2.merge(marketInvColumn, how='left', left_on='Sku', right_on='Sku')
		merged2 = merged2.rename(columns={'Units_y': (str(eachMarket) + '\\Covered\\Stores\\Units'), 'Units_x':'Units'})

		
		
		######  DC/DDC Units Column For Each Store   ######
		#sum DC/DDC inventory for each sku based on what DC/DDC's match up to each store
		#remove duplicate skus / storeId's from marketdf to create a list of dc/ddc per store
		listofDcDDC = merged2 [['StoreId','DC Store#','DDC Store#']]
		listofDcDDC = listofDcDDC.drop_duplicates('StoreId')
		
		#create a list of storeid / dc / ddc for each store in marketdf
		listofDcDDC = listofDcDDC[['StoreId','DC Store#','DDC Store#']].values
		listofDcDDC = listofDcDDC.tolist()
		
		allstoreDcDDCList = []
		for eachStore in listofDcDDC:
			#identify dc and ddc store# that matches to each store in market
			thisStore = eachStore[0]
			dcstore = eachStore[1]
			ddcstore = eachStore[2]			
			#for each group of dc/ddc in lists, filter df by dc/ddc storeid's and create new storeDcDDCInv df
			storeDcDDCInv1 = whseInvDf [ whseInvDf['StoreId'] == dcstore ]
			storeDcDDCInv2 = whseInvDf [ whseInvDf['StoreId'] == ddcstore ]
			storeDcDDCInv = pd.concat([storeDcDDCInv1, storeDcDDCInv2])
			
			#sum this store's dc/ddc units per sku and turn series into df
			storeDcDDCInv = storeDcDDCInv.groupby(by=['Sku'])['Inventory'].sum()
			storeDcDDCInv = pd.DataFrame({'zDC/DDC Units':storeDcDDCInv.values,'Sku':storeDcDDCInv.index})
			
			#save this store's number of dc/ddc units to a list, 
			#then add this list to allstoresDcDDClist to hold numbers of all skus for all stores in market
			storeDcDDCList = storeDcDDCInv [['Sku', 'zDC/DDC Units']]
			storeDcDDCList = storeDcDDCList.values.tolist()
			for eachList in storeDcDDCList:
				eachList.append(thisStore)
				allstoreDcDDCList.append(eachList)
			
		#turn allstoreDcDDCList into a new df
		allstoreDcDDCInv = pd.DataFrame(allstoreDcDDCList, columns=['Sku','zDC/DDC Units','StoreId'])
		
		#sum the total inventory for dc/ddc
		allstoreDcDDCInv = allstoreDcDDCInv.groupby(by=['Sku', 'StoreId'])['zDC/DDC Units'].sum()	
		
		#allstoreDcDDCInv = pd.DataFrame({'Units':allstoreDcDDCInv.values,'Sku':allstoreDcDDCInv.index})
		allstoreDcDDCInvDf = pd.DataFrame(allstoreDcDDCInv)
		allstoreDcDDCInvDf = allstoreDcDDCInvDf.reset_index()
		
		#add column for storeDcDDCInv to marketdf
		merged4 = merged2.merge(allstoreDcDDCInvDf, how='left', left_on=['Sku', 'StoreId'], right_on=['Sku', 'StoreId'])

				

		######  Pivot Table Section  ######
		#rename Units - starting with "a" to pull into pivotdf in correct order before zDC/DDC Units	
		merged4 = merged4.rename(columns={'Units':'aUnits'})
		merged4 = merged4.fillna(value=0)
		
		#turning full merged4 df into pivot table results in a memory error
		#create 2 new dataframes for skus having 0 & >0 total units in market to create 2 pivot tables, then combine them into one
		merged5 = merged4 [ merged4[(str(eachMarket) + '\\Covered\\Stores\\Units')] > 0 ]
		merged6 = merged4 [ merged4[(str(eachMarket) + '\\Covered\\Stores\\Units')] == 0 ]

		pivotindex = ['Region','Market','Sku', 'Model', 'Size','National DC/DDC Inventory',
			(str(eachMarket) + '\\Covered\\Stores\\Units')]

		#market units > 0 sorted by model / size (descending)			
		pivotdf1 = pd.pivot_table(merged5, values=['zDC/DDC Units','aUnits'], index=pivotindex, columns='StoreIdMerged', fill_value=0, dropna=True).reset_index()
		pivotdf1 = pivotdf1.sort_values(by=['Model', 'Size'], ascending=[0,0])	
		
		#market units = 0 sorted by model / size (descending)			
		pivotdf2 = pd.pivot_table(merged6, values=['zDC/DDC Units','aUnits'], index=pivotindex, columns='StoreIdMerged', fill_value=0, dropna=True).reset_index()
		pivotdf2 = pivotdf2.sort_values(by=['Model', 'Size'], ascending=[0,0])
		
		#combine pivots into one
		pivotdf = pivotdf1.append(pivotdf2, ignore_index=False)
		
		#reorder column labels moving storeId to row 1
		pivotdf = pivotdf.reorder_levels([1, 0], axis=1)
		pivotdf = pivotdf.sort_index(level=[1],axis=1, ascending=[False], na_position='first')

		#rename Units & DC/DDC Units columns to remove first character that was used for sorting
		pivotdf = pivotdf.rename(columns={'zDC/DDC Units': 'DCs', 'aUnits': 'Units'})

		
		
		################    TEMP FILE SECTION    ################
		#write pivot df to temp excel doc
		tempWbSh.range('A1').value = pivotdf
		
		#Set variables for region and market, then delete columns
		market = tempWbSh.range("C5").value
		region = tempWbSh.range("B5").value
		tempWbSh.range("A:A").api.EntireColumn.Clear()
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
		
		#create 2 new rows with store# and tier, then create list from each store column used to sort stores left-to-right
		tempWbSh.range("1:2").api.EntireRow.Insert()
		storeTierRange = tempWbSh.range("G5:AP5")
		for tier in storeTierRange:
			storeTierValue = tempWbSh.range(tier).value
			if storeTierValue is not None:
				storeTierColumn = tier.column					
				storeTier = tempWbSh.range(tier).value
				tempWbSh.range(2, storeTierColumn).value = storeTier[5:6]

		storeIdRange = tempWbSh.range("G3:AP3")
		storeList = []
		for store in storeIdRange:
			storeIdValue = tempWbSh.range(store).value
			if storeIdValue is not None:
				storeIdColumn = store.column
				storeIdNumber = tempWbSh.range(store).value					
				tempWbSh.range(1, storeIdColumn).value = storeIdNumber[4:len(storeIdNumber)]
				storeList.append(tempWbSh.range((1, storeIdColumn),(120,storeIdColumn)).value)
		
		#sort columns by by tier (ascending), then storeId (ascending), then Units/DC's label (descending)
		storeList = sorted(storeList, key=itemgetter(1,0,-5))

		#write the sorted store columns (G-AP) to the file
		tempWbSh.range("G1:AP101").api.EntireColumn.Clear()
		for n in range(0,len(storeList)):
			r = tempWbSh.range((1, n+7), (120, n+7))
			r.options(transpose = True).value = storeList[n]

					
		###### FINAL MARKET WORKBOOK SECTION ######
		#create a new workbook from the template tab, rename sheet as "TV Inventory", delete unused Sheet1
		new_wb = xlwings.Book()
		newWbSh = new_wb.sheets[0]
		valuedWb.sheets[templateWs].api.Copy(Before=newWbSh.api)
		new_wb.sheets[0].name = 'TV Inventory'
		new_wb.sheets['Sheet1'].delete()
		
		#copy data from temp file to new workbook, then set freeze panes
		marketValuesTemp = tempWbSh.range('A3:AP130').value 
		new_wb.sheets['TV Inventory'].range('A1:AP127').value = marketValuesTemp
		new_wb.FreezePanes = False
		new_wb.sheets['TV Inventory'].range("G5").select()
		new_wb.FreezePanes = True
		
		#save new workbook as "[market] BBY TV Inventory.xlsm" to folder: C:\Users\username\Documents\BBY TV Inventory Market Reports\
		new_wb.save(r'C:\\Users\\'+ homedir + '\\Documents\\BBY TV Inventory Market Reports\\' + str(eachMarket) + ' BBY TV Inventory.xlsx')
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
#excel.ScreenUpdating = True
os.startfile(filename)
