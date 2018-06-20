'''using pyhon 2.7'''
import re, os
import openpyxl
from pandas import DataFrame
# import PyPDF2

resources = "C:\\Users\\Etienne\\Documents\\GitHub\\From-PDF-to-Excel\\ressources"
excel_output = "C:\\Users\\Etienne\\Documents\\GitHub\\From-PDF-to-Excel\\excel_output"

PyPDF2_path = "C:\\PYTHONLIBS\\PyPDF2-1.26.0"


import sys
sys.path.append(PyPDF2_path)
import PyPDF2

def grapLastPagePDF(path, fileName):
	'''
	from one PDF, grab the last page and 
	extract text from it
	'''
	pdfFileObj = open(path, 'rb')
	pdfReader = PyPDF2.PdfFileReader(pdfFileObj) 	# creating a pdf reader object
	numbPage = pdfReader.getNumPages()				# get the number of pages

	try:
		pageObj = pdfReader.getPage(numbPage-1) 	# creating a page object and read last page
		textinpage = pageObj.extractText() 			# extracting text from page
	except:
		print("failed due to encryption:", fileName)
		textinpage = [False]
	
	pdfFileObj.close()
	return textinpage

def transformToDf(text):
	'''
	transform text into a df with 
	regular expression
	'''
	# TO DO: regular expression to grab the list of item
	t = ["Total Revenue","Cost of Goods Sold","Gross Profit","Operating Expenses", "Salaries",
		"Rent", "Utilities", "Depreciation","Total Operating Expenses","Operating Profit (EBIT)",
		"Interest Expense","Income before taxes (EBT)","Taxes","Net Income","Number of Shares Outstanding","Earnings Per Share (EPS)"]

	# list that will contain the extracted items
	n = []

	for i in t:
		# 1.remove special char of the list
		item = i.replace(' ', '[ ]').replace('(', '[(]').replace(')', '[)]')

		# 2.regular exp to grab the figures
		alllist = re.compile(r'''
			%s(			#item to extract 
			(\s)+? 		#space before number - 100 000.00
			([-]\s+)?	#optional negative sign
			(\d+[ ])?
			(\d+[ ])?	#100 and 1 space
			\d+[.]\d+	#000 and . and 00
			)''' %item , re.VERBOSE)

		listofalllist = alllist.findall(text)
		
		# 3.append to the list n
		n.append([i[0].replace('\n','').replace(' ','')  for i in listofalllist])

	dictionary = dict(zip(t, n)) #create a dictionnary out of t and n
	return DataFrame(dictionary)	#create a df out of t and n

def loopAllPDF(PDFdirectory):
	'''
	Loop through all the pdf file of a folder
	merge them and parse into excel
	'''
	alldf = DataFrame()
	for fileName in os.listdir(PDFdirectory):
		if fileName.endswith(".pdf"):
			df = transformToDf(grapLastPagePDF(PDFdirectory+'\\'+fileName, fileName))
			alldf = alldf.append(df)
			print("file extracted:", fileName)

	alldf.to_excel(excel_output + "\\" + 'allPDF_py2.7.xlsx', sheet_name='sheet1', index=False)

loopAllPDF(resources)


