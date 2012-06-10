from Excel_Extractor import *;
from MySQL_Importer import *;

#main
#excel extraction
temp=Excel_Extractor("Dataentrysheet.xls");
temp2=MySQL_Importer("localhost","root","password","databasename");  #HOME

choice=raw_input("\n\nDo u want to import Sheet1 data? (y/n)");
if choice=="Y" or choice=="y":
	print "Importing ITEM data...";
	masterList=temp.extractSheetData_Records(0); #extracting item data
	temp2.batch_insertQuery_Item(masterList);#dump into mysql

print "====Finished====";

