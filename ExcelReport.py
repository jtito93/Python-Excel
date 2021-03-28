import pandas as pd
import xlsxwriter
import numpy as np

class Excel:
	
	def __init__(self, dataframe, file_name, sheet_name):
		self.dataframe = dataframe
		self.file_name = file_name
		self.writer = pd.ExcelWriter(file_name, engine="xlsxwriter")
		self.dataframe.to_excel(self.writer, sheet_name = sheet_name, startrow=1, header=False, index=False)
		self.workbook = self.writer.book
		
		# Header format
		self.worksheet = self.writer.sheets[sheet_name]
		self.set_header_format(True, 'center', '#D7E4BC', 1, 'black')
		self.header_format = self.workbook.add_format(self.get_header_format())		
		
		# Write the column headers with the defined format.
		for col_num, value in enumerate(self.dataframe.columns.values):
			self.worksheet.write(0, col_num, value, self.header_format)
	
	def set_header_format(self, bold, align, bg_color, border, font_color):
		self.header_bold = bold
		self.header_align = align
		self.header_bg_color = bg_color
		self.header_border = border
		self.header_font_color = font_color
	
	def get_header_format(self):
		return{
			'bold': self.header_bold,
			'valign': self.header_align,
			'fg_color': self.header_bg_color,
			'border': self.header_border,
			'font_color': self.header_font_color
		}
	
	def update_worksheet_header(self, dataframe, sheet_name):
		# Header format
		self.worksheet = self.writer.sheets[sheet_name]
		self.dataframe = dataframe
		self.header_format = self.workbook.add_format(self.get_header_format())		
		
		# Write the column headers with the defined format.
		for col_num, value in enumerate(self.dataframe.columns.values):
			self.worksheet.write(0, col_num, value, self.header_format)
		
	def add_worksheet(self, dataframe, sheet_name):
		self.dataframe = dataframe
		self.dataframe.to_excel(self.writer, sheet_name = sheet_name, startrow=1, header=False, index=False)
	
	def set_column_format(self, dataframe, sheet_name, column_index, format_type):
		worksheet_format = self.workbook.add_format(format_type)
		self.worksheet = self.writer.sheets[sheet_name]
		self.worksheet.set_column(column_index, column_index, None, worksheet_format)
		
	def adjust_column_size(self, dataframe, sheet_name, include_header=True):
		self.worksheet = self.writer.sheets[sheet_name]
		measurer = np.vectorize(len)
		columns = measurer(dataframe.values.astype(str)).max(axis=0)
		for index, size in enumerate(columns):
			if include_header:
				header_size = len(dataframe.columns[index])
				new_size = header_size if header_size > size else size
				self.worksheet.set_column(index, index, new_size) 
			else:
				self.worksheet.set_column(index, index, size) 	
		
	def save(self):
		self.writer.save()



def start():
		
	file_name = "Sample.csv"
	excel_file = "Sample.xlsx"

	df = pd.read_csv(file_name, sep=',')
	
	print('Creating Excel file...')
	
	my_excel = Excel(df, excel_file, 'Sheet1')	
	my_excel.set_header_format(1, 'left', '#4293f5', 1, 'white')
	my_excel.update_worksheet_header(df, 'Sheet1')
	my_excel.set_column_format(df, 'Sheet1', 2, {'num_format': '$#,##0.00'})
	my_excel.set_column_format(df, 'Sheet1', 6, {'num_format': '%0'})
	my_excel.adjust_column_size(df, 'Sheet1')
	my_excel.save()

	print('File created!')

if __name__ == '__main__':
	start()
	
	
