import _mysql

class MySQL_Importer():
	def __init__(self,host,username,pwd,dbname):
		#connect to db
		self.con=_mysql.connect(host,username,pwd,dbname)
		print "Connected to DB : ",dbname;
	
	"""Insert record into Recipe table"""
	def __insertQuery_Recipe(self,recipe_name,food_category):
		#store in dictionary
		dictionary={'recipe_name':recipe_name,'food_category':food_category};
		#form the query string
		str="INSERT INTO RECIPE(RECIPE_NAME,FOOD_CATEGORY) VALUES(\'%(recipe_name)s \', \'%(food_category)s\')" % dictionary ;
		print "Insert Query:  " ,str;
		
		#insert into table
		c=self.con.query(str);

	def __getSingleValue(self,queryString):
		#perform query select operation
		self.con.query(queryString);
		result=self.con.store_result();
		count=result.num_rows();
		#print "Printing the result: Size=",count;
		if(count==0):
			return "nothin";
		else:
			if(count==1):
				tempTuple = result.fetch_row();#retrieves a single record (but it is stored in 2 layer tuple)
				for x in tempTuple:
					for y in x:
						return str(y);  #printing each column value
			else:
				return "More than one value found during getSingleValue()= ",queryString;
			
	def displayResult(self):
		#perform query select operation
		self.con.query("SELECT recipe_name from RECIPE");
		result=self.con.store_result();
		count=result.num_rows();
		print "Printing the result: Size=",count;
		while count>0:
			tempTuple = result.fetch_row();#retrieves a single record (but it is stored in 2 layer tuple)
			for x in tempTuple:
				for y in x:
					print y,  #printing each column value
			count=count-1;
			print ""

