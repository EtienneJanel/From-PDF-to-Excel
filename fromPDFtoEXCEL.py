import PyPDF2, re, os
from pandas import DataFrame


def grapLastPagePDF(path):
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
		print("failed due to encryption")
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
	for listePDF in os.listdir(PDFdirectory):
		if listePDF.endswith(".pdf"):
			df = transformToDf(grapLastPagePDF(PDFdirectory+'\\'+listePDF))
			alldf = alldf.append(df)

	alldf.to_excel('allPDF.xlsx', sheet_name='sheet1', index=False)

df = DataFrame()

PDFdirectory = "C:\\Users\\Etienne\\Documents\\GitHub\\From-PDF-to-Excel\\PDF-to-copy"

# loop through all PDF
loopAllPDF(PDFdirectory)

# To keep
# number = re.compile(r'(([-]\s+)?(\d+[ ])?\d+[.]\d+)')	# regular expression of the numer such as -123 456.00
# grossprofit = re.compile(r'Gross Profit((\s)+?([-]\s+)?(\d+[ ])?\d+[.]\d+)') # regular expression of the gross profit
# df = df.append(df2)	#append another dataframe
# current_directory = os.getcwd() #current working directory

# text = grapLastPagePDF(path)	# content of the last page
# df = transformToDf(text)		# to a df
# df.to_excel('from_PDF_to_XLSX.xlsx', sheet_name='sheet1', index=False) #parse into excel

## list of file in folder
# listePDF = os.listdir(PDFdirectory)
# print("liste of PDF", listePDF , "\n")
