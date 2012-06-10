from sets import Set
import xlrd

class Excel_Extractor():
	def __init__(self,filename):
		self.book = xlrd.open_workbook(filename);
		print "Loaded Excel-sheet : ",filename;
		print "Details:";
		print "Sheet names:",self.book.sheet_names();

	"""Displays the sheet structure"""
	def displaySheetInfo(self,sheetIndex):
		sheet=self.book.sheet_by_index(sheetIndex);        
		print "\tDetails:";
		print "\t*Sheet structure:"
		print "\t=================";
		print "\tNo.of rows : ",sheet.nrows;
		print "\tNo.of cols : ",sheet.ncols;
		print "\t=================";

	"""Extracts the column names from the given sheet which is later used as table cols for insertion.
		Always scans the 1st row only!"""
	def extractSheetData_ColNames(self,sheetIndex):
		colMap={};
		sheet=self.book.sheet_by_index(sheetIndex);        
		#create a dictionary to store the current record data(names are based on table cols)
		for y in range(0,sheet.ncols): #for each column
			colMap.update({y:sheet.cell_value(0,y)});
		print "Columns from Sheet[name,index]=[",self.extractSheetData_TableName(sheetIndex),",",sheetIndex,"]";
		return colMap;

	"""Extracts the table name from the given sheet"""
	def extractSheetData_TableName(self,sheetIndex):
		temp=self.book.sheet_names();
		return temp[sheetIndex];

	"""Extracts the spreadsheet data and returns in list-in-list form"""
	def extractSheetData_Records(self,sheetIndex,row_offset=2,col_offset=0):
		masterList=[];#this List stores all the records(dictionary form).
		sheet=self.book.sheet_by_index(sheetIndex);        
		colMap=self.extractSheetData_ColNames(sheetIndex);#extracting the column-names
		maxrows=sheet.nrows;
		maxcols=sheet.ncols;
		#creating the childMap from the colMap
		childMap={};
		for x in colMap:
			childMap.update({colMap[x]:""});
		print "Extracting data from Sheet:",sheetIndex;
		for x in range(row_offset,maxrows):
			#getting a copy of childMap(names are based on table cols)
			temp_childMap=childMap.copy();
			for y in range(col_offset,maxcols): #for each column
				#storing the data in childList.Its index is calculated from col-index which is mapped in 'listStructure'
				temp_childMap[colMap[y]]=sheet.cell_value(x,y);
				#print sheet.cell_value(x,y);
			masterList.append(temp_childMap);#adding to masterList 
			#print "============";
		print "  MasterList Size : ",len(masterList);
		return masterList;
		
	def getNonexistRecordInfo(self,master_sheetIndex,master_column_no,slave_sheetIndex,slave_column_no,row_offset=2):
		masterList=[];
		slaveList=[];
		print "**  Validating data between Sheets[",master_sheetIndex,"] and [",slave_sheetIndex,"]  **";
		#Creating masterData list
		sheet=self.book.sheet_by_index(master_sheetIndex);        
		maxrows=sheet.nrows;
		print "\nExtracting masterData from Sheet : ",master_sheetIndex," at COLUMN : ",master_column_no;
		for x in range(row_offset,maxrows):
			masterList.append(str(sheet.cell_value(x,master_column_no)));
			
		print "  MasterData list         COUNT  : ",len(masterList);
		print "  Removing duplicate data COUNT  : ",len(set(masterList));
		masterList=list(set(masterList));#converting set back to list
		
		#Creating slaveData list
		sheet=self.book.sheet_by_index(slave_sheetIndex);        
		maxrows=sheet.nrows;
		print "\nExtracting slaveData from Sheet  : ",slave_sheetIndex," at COLUMN : ",slave_column_no;		
		for x in range(row_offset,maxrows):
			slaveList.append(str(sheet.cell_value(x,slave_column_no)));

		print "  SlaveData list          COUNT  : ",len(slaveList);
		print "  Removing duplicate data COUNT  : ",len(set(slaveList));
		slaveList=list(set(slaveList));#converting set back to list

		print "Master List\n===========";
		
		for temp in masterList:
			print temp;
		
		print "Slave List\n===========";
		
		for temp in slaveList:
			print temp;
		
		print "\n==Now showing missing slaveData in masterList==";
		print "---------------"
		for data in slaveList:
			if data not in masterList:
				print data;
		print "---------------"				
		print "=== Non-existing record finding done ===";
		
	def getDuplicateRecordInfo(self,sheetIndex,column_no,row_offset=2):
		datalist={};#this List stores all the records(dictionary form).
		
		sheet=self.book.sheet_by_index(sheetIndex);        
		maxrows=sheet.nrows;

		print "Extracting data from Sheet:",sheetIndex;
		for x in range(row_offset,maxrows):
			data=str(sheet.cell_value(x,column_no));
			#print "given data : ",data;
			#checking whether the key already exists or not
			if(datalist.has_key(data)):
				print data,"   key found!..updating";
				datalist[data].append(x+1); #adding +1 for equalizing to actual row-index
			else:			#if a key already exits for this data..
				print data,"   key not found..adding";
				datalist[data]=[x+1];
			
		print "\n\n\n**** Now showing only duplicate data information for sheet:",sheetIndex," ****"
		#now,the keys having the values more than 1 are printed 
		
		for temp in datalist:
			if(len(datalist[temp]) > 1):
				print "Data : ",temp,"		Row-Indexes : ",datalist[temp];
				
		print "==end of duplicate checking==";
		
	"""	
	#Returns dictionary containing table cols(mapped to empty values)
	def getProcessedColDictionary(self,sheetIndex):
		colDict={};
		#get the colNames list
		colList=self.extractSheetData_ColNames(sheetIndex);
		for x in colList:
			colDict.update({x:"test"});
		return colDict;
	"""

def UserMenu():
	print "Menu:";
	print "-----";
	print "1.Want to find duplicates items?";
	print "2.Want to validate recipe and ingredient datasheet?";
	print "0.Nothing";
	return raw_input("\n\nChoice : ");
	
#Main
temp=Excel_Extractor("Dataentrysheet.xls");
print "\n\nExcel2MySQL Extractor";
print "--------------";

choice=UserMenu();

while choice!="0":
	if choice=="1":
		temp.getDuplicateRecordInfo(0,0);#sheet_number,col_index
	elif choice=="2":
		temp.getNonexistRecordInfo(master_sheetIndex=1,master_column_no=0,slave_sheetIndex=2,slave_column_no=0);
		
	temp_choice=raw_input("\n\nWant to do another operation? (y/n)")
	if temp_choice=="Y" or temp_choice=="y":
		choice=UserMenu();
	else:
		choice="0";

print "....end :-)";
