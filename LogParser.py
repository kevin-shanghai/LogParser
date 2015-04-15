#/usr/bin/python
import os
import sys
try:
	import xlwt as ExcelWrite
except ImportError:
	print"Please install the third party module xlwt first....."
	sys.exit()

class LogParser:
	Dose1AllLogParseValue = [] 
	Dose2AllLogParseValue = []
	def __init__(self):
		self.DOSE1_TAG = "dose1:iDoseADC0Value is:"
		self.DOSE2_TAG = "dose2:iDoseADC0Value is:"
		self.DOSE_VALUE_SPLITER = "###"

	def getLogFiles(self):
		allFiles = os.listdir(os.getcwd())
		allLogFiles = [singleFile for singleFile in allFiles if singleFile.endswith("uihlog")]
		Logprefix = [int(x.split(".")[0]) for x in allLogFiles ]
		Logprefix.sort()
		allSortedLogFiles = [str(x)+".uihlog" for x in Logprefix]
		if len(allSortedLogFiles) == 0:
			raise AssertionError("log format must be *.uihlog, please check!")
		return allSortedLogFiles 

	def parseSingleLogFile(self, logfile):
		print logfile
		f = open(logfile)
		for line in f.readlines():
			dose1LogInSingleFile = []
			dose2LogInSingleFile = []
			if self.DOSE1_TAG in line and self.DOSE_VALUE_SPLITER in line:
				dose1LogInSingleFile.append("Dose1:")
				splitdoseValue = line.split(self.DOSE_VALUE_SPLITER)
				try:
					dose1LogInSingleFile.append(splitdoseValue[9])
					dose1LogInSingleFile.append(splitdoseValue[1])
					dose1LogInSingleFile.append(splitdoseValue[3])
					dose1LogInSingleFile.append(splitdoseValue[5])
					dose1LogInSingleFile.append(splitdoseValue[7])
				except RuntimeError:
					print"Please check the log format, something error...."
			elif self.DOSE2_TAG in line and self.DOSE_VALUE_SPLITER in line:
				dose2LogInSingleFile.append("Dose2:")
				splitdoseValue = line.split(self.DOSE_VALUE_SPLITER)
				try:
					dose2LogInSingleFile.append(splitdoseValue[9])
					dose2LogInSingleFile.append(splitdoseValue[1])
					dose2LogInSingleFile.append(splitdoseValue[3])
					dose2LogInSingleFile.append(splitdoseValue[5])
					dose2LogInSingleFile.append(splitdoseValue[7])
				except RuntimeError:
					print"Please check the log format, something error...."
			else:
				pass
			if len(dose1LogInSingleFile):
				LogParser.Dose1AllLogParseValue.append(dose1LogInSingleFile)
			if len(dose2LogInSingleFile):
				LogParser.Dose2AllLogParseValue.append(dose2LogInSingleFile)
		f.close()


	def parseAllLogFiles(self):
		for logFile in self.getLogFiles():
			self.parseSingleLogFile(logFile)

	def writeParsedLogToFiles(self):
		file_dose1_output = open(os.path.abspath('.')+'/dose1_output.txt','w') 
		file_dose2_output = open(os.path.abspath('.')+'/dose2_output.txt','w') 
		for item in LogParser.Dose1AllLogParseValue:
			file_dose1_output.write("\t".join(item) + "\n")
		for item1 in LogParser.Dose2AllLogParseValue:
			file_dose2_output.write("\t".join(item1) + "\n")
		file_dose1_output.close()
		file_dose2_output.close()

	def getSingleSheetWriteContentDic(self, all_sheets):
		singleSheetDose1WriteContentDic = {}
		singleSheetDose2WriteContentDic = {}
		print "in getSingleSheetWriteContentDic, WriteExcel.excel_sheet_nums is:", WriteExcel.excel_sheet_nums
		for i in range(WriteExcel.excel_sheet_nums):
			singleSheetDose1WriteContentDic[all_sheets[i]] = LogParser.Dose1AllLogParseValue[i*60000:(i+1)*60000]
		if WriteExcel.excel_sheet_nums > 1:
			singleSheetDose1WriteContentDic[all_sheets[-1]] = LogParser.Dose1AllLogParseValue[(WriteExcel.excel_sheet_nums-1) * 60000:]

		for i in range(WriteExcel.excel_sheet_nums):
			singleSheetDose2WriteContentDic[all_sheets[i]] = LogParser.Dose2AllLogParseValue[i*60000:(i+1)*60000]
		if WriteExcel.excel_sheet_nums > 1:
			singleSheetDose2WriteContentDic[all_sheets[-1]] = LogParser.Dose2AllLogParseValue[(WriteExcel.excel_sheet_nums-1) * 60000:]

		return singleSheetDose1WriteContentDic, singleSheetDose2WriteContentDic

	

class WriteExcel:
	excel_sheet_nums = 0
	def __init__(self):
		self.title = ["Device", "Test Num", "ADC0", "Seg12", "Seg34", "ADCBakup"]

	def getTitleStyle(self):
		font0= ExcelWrite.Font()
		font0.name= 'Times New Roman'
		font0.colour_index= 2
		font0.size = 20
		font0.bold= True
		style0= ExcelWrite.XFStyle()
		style0.font= font0
		return style0

	def writeExcelTitle(self, sheet):
		for i in range(len(self.title)):
			sheet.write(0, i, self.title[i], self.getTitleStyle())
			sheet.write(0, i+8, self.title[i], self.getTitleStyle())

	def getExcelSheetNums(self):
		sheetNums = len(LogParser.Dose1AllLogParseValue) / 60000 + 1
		WriteExcel.excel_sheet_nums = sheetNums
		return sheetNums

	def createExcelSheets(self, book):
		all_sheets = []
		for i in range(1, self.getExcelSheetNums() + 1):
			all_sheets.append(book.add_sheet("Sheet" + str(i)))
		return all_sheets

	def writeEverySheet(self, all_sheets, dose1_write_content_sheet_dic, dose2_write_content_sheet_dic):
		colum_nums = len(LogParser.Dose1AllLogParseValue[0])
		for i in range(len(dose1_write_content_sheet_dic)):
			single_sheet_row_num = len(dose1_write_content_sheet_dic[all_sheets[i]])
			for row in xrange(single_sheet_row_num):
				for colum in xrange(colum_nums):
					cell_content_write = dose1_write_content_sheet_dic[all_sheets[i]][row][colum]
					all_sheets[i].write(row+1, colum, cell_content_write)

		for i in range(len(dose2_write_content_sheet_dic)):
			single_sheet_row_num = len(dose2_write_content_sheet_dic[all_sheets[i]])
			for row in xrange(single_sheet_row_num):
				for colum in xrange(colum_nums):
					cell_content_write = dose2_write_content_sheet_dic[all_sheets[i]][row][colum]
					all_sheets[i].write(row+1, colum + 8, cell_content_write)


	def write2Excel(self, parser):
		book = ExcelWrite.Workbook()
		dose1_row_nums = len(LogParser.Dose1AllLogParseValue)
		dose1_colum_nums = len(LogParser.Dose1AllLogParseValue[0])
		print "dose1_row_nums:",dose1_row_nums

		all_sheets = self.createExcelSheets(book)
		for i in range(self.getExcelSheetNums()):
			self.writeExcelTitle(all_sheets[i])	
		dose1_write_content_sheet_dic, dose2_write_content_sheet_dic = parser.getSingleSheetWriteContentDic(all_sheets)
		self.writeEverySheet(all_sheets, dose1_write_content_sheet_dic, dose2_write_content_sheet_dic);
		book.save('DoseLog.xls')

if __name__ == '__main__':
	logparser = LogParser()
	print logparser.getLogFiles()
	logparser.parseAllLogFiles()
	# print LogParser.Dose1AllLogParseValue
	excelWriter = WriteExcel()
	excelWriter.write2Excel(logparser)


