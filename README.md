import pandas as pd
from openpyxl import load_workbook


WORKSHEET = "D:\Py\Book1.xlsx"
SHEETS = ["Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5"]
MSHEET = "MasterSheet"
workbook=load_workbook(r"D:\Py\Book1.xlsx")

class Aggregator:
	def __init__(self, worksheet, sheets):
		self.worksheet, self.sheets = worksheet, sheets
		self.dfs = pd.read_excel(worksheet, sheet_name=sheets, parse_dates=False)

	def get_input(self, c=0):
		query = input("Enter PS Number / Email / Name: ")
		if query and c < 3:
			try:
				query = int(query)
				searchid = "Ps No"
			except ValueError:
				if "@" in query:
					searchid = "Email"
				else:
					searchid = "Name"
			return query, searchid
		elif c == 3:
			print("Too many wrong attempts, try again later.")
			exit()
		else:
			print("No input found, try again!")
			return self.get_input(c + 1)

	def search(self, query, searchid):
		print(f"Searching {searchid.lower()} `{query}`...")
		fields = {}
		for x in self.dfs.values():
			fields.update(x[x[searchid] == query].to_dict(orient="list"))
		# print(fields)

		if fields[searchid]:
			print("Found.")
			return pd.DataFrame.from_dict(fields)

		print(f"Couldn't find the {searchid.lower()} in sheets!")
		exit()

	def add_to_master(self, df):

		#df["Entry Time"] = df["Entry Time"].dt.strftime("%I:%S %p")
		# df["Exit Time"] = df["Exit Time"].dt.strftime("%H:%S %p")
		# df["Start Date"] = df["Start Date"].dt.strftime("%d/%m/%Y")
		# df["End Date"] = df["End Date"].dt.strftime("%d/%m/%Y")

		book = load_workbook(self.worksheet)   # loading the workbook
		with pd.ExcelWriter(self.worksheet, mode="a") as writer:  # opening a Excel Writer instance
			writer.book = book  # changing the workbook of the writer to our current workbook
			writer.sheets = {ws.title: ws for ws in book.worksheets}  # adding the worksheets to it
			# getting the last row if not found set it to 0
			try:
				startrow = writer.sheets[MSHEET].max_row
			except KeyError:
				startrow = 0
			# checking if we have already have a master sheet
			if MSHEET in writer.book.sheetnames:
				df.to_excel(
					writer,
					index=False,
					header=False,
					sheet_name=MSHEET,
					startrow=startrow,
				)
			else:
				df.to_excel(writer, index=False, sheet_name=MSHEET, startrow=startrow)
		print(f"Added to {MSHEET}.")


if __name__ == "__main__":

	agg = Aggregator(WORKSHEET, SHEETS)

	query, searchid = agg.get_input()

	df_m = agg.search(query, searchid)

	agg.add_to_master(df_m)



#=======================
# Creating another mastersheet with total number of rows and columns.
if 'Mastersheet1' in SHEETS:
    mas1 = workbook['Mastersheet1']
# If mastersheet is there, it will remove.    
    workbook.remove(mas1)   
# It will again create a new mastersheet.                                
workbook.create_sheet("Mastersheet1")

 

# Getting maximum row count
row_num=workbook['MasterSheet'].max_row
# Getting maximum column count
column2=workbook['MasterSheet'].max_column

mas1=workbook['Mastersheet1']
# Printing total number of rows in mastersheet 2
mas1.cell(row=1,column=1).value="Total number of rows"
# Printing total number of columns in mastersheet 2
mas1.cell(row=3,column=1).value="Total number of columns"
# Getting row value 
mas1.cell(row=1,column=2).value=row_num
# Getting column value
mas1.cell(row=3,column=2).value=column2
# Saving the file.
workbook.save(r"D:\Py\Book1.xlsx")
