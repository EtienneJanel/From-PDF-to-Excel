
import PyPDF2, re, os
from pandas import DataFrame


df = DataFrame()
directory = "C:\\Users\\Etienne\\Documents\\GitHub\\From-PDF-to-Excel"
filename = "pnl.pdf"
path = directory + filename

def grapLastPagePDF( path):
	pdfFileObj = open(filename, 'rb') 	
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

# TO DO: modify the Loop-all-PDF to match the new df
def loopAllPDF(directory):
	for filename in os.listdir(directory):
		if filename.endswith(".pdf"):
			graplastpagePDF(df, filename)

			continue
		else:

			continue

	df.to_excel('fromPDFtoXLSX.xlsx', sheet_name='sheet1', index=False)
	pass

text = grapLastPagePDF(path)	# content of the last page

# number = re.compile(r'(([-]\s+)?(\d+[ ])?\d+[.]\d+)')	# regular expression of the numer such as -123 456.00
# grossprofit = re.compile(r'Gross Profit((\s)+?([-]\s+)?(\d+[ ])?\d+[.]\d+)') # regular expression of the gross profit

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
		(\d+[ ])?	#100 and 1 space
		\d+[.]\d+	#000 and . and 00
		)''' %item , re.VERBOSE)

	listofalllist = alllist.findall(text)
	
	# 3.append to the list n
	n.append([i[0].replace('\n','').replace(' ','')  for i in listofalllist])

dictionary = dict(zip(t, n)) #create a dictionnary out of t and n
df = DataFrame(dictionary)	#create a df out of t and n

# df = df.append(df2)	#append another dataframe

df.to_excel('from_PDF_to_XLSX.xlsx', sheet_name='sheet1', index=False) #parse into excel



