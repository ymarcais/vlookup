import pandas as pd
import openpyxl


class Excel_():

	def __init__(self) -> None:
		pass

	#Research elements from 1st column of file2 in the first column of file 1.
	#If true set file2 element row in new file
	def vlookup(self, path1, path2):

		#Load files into pandas dataframe
		df1 = pd.read_excel(path1, sheet_name="Sheet1")
		df2 = pd.read_excel(path2, sheet_name="Sheet2")
		
		#Extract value of both first columns
		value_sheet1 = df1.iloc[:, 0]
		value_sheet2 = df2.iloc[:, 0]

		#Filter rows from df2 where the value in the first column is in df1
		match_row = df2[value_sheet2.isin(value_sheet1)]

		#Save the matching rows to a new Excel file in xls
		match_row.to_excel('matching_rows.xlsx', index=False)


def main():
	path1 = "file1.xls"
	path2= "file2.xls"

	xcl = Excel_()
	try:
		xcl.vlookup(path1, path2)
		print("vlookup sucess process")
	except ValueError:
		print("Oops error vlookup")

if __name__ == "__main__":
	main()


