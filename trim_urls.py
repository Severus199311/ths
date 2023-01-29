#把excel表格中url不需要的部分截掉。
#只涉及excel，与种子库、爬虫什么的无关

import xlwt, xlrd
from xlutils.copy import copy
import time, re
from logger import logger

class TrimUrls():

	def __init__(self):

		self.workbook_title = 'urls_to_be_trimmed.xls'
		self.sheet_name = 'Sheet1' #待录入种子所在的表名
		self.sheet_order = 0 #待录入种子所在的表序号

		self.url_trimmed_coloumn = 7

		self.workbook_read = xlrd.open_workbook(self.workbook_title) #打开excel文件，用于读取数据
		self.sheet_data = self.workbook_read.sheet_by_name(self.sheet_name) #读取表数据
		self.keys_list = self.sheet_data.row_values(0) #获取列名，格式为list
		self.workbook_write = copy(self.workbook_read) #复制excel文件，用于写入
		
		self.processing_row = 0 #初始行序号为0

	def fetch_url_undone(self):
		self.processing_row += 1
		value_list = self.sheet_data.row_values(self.processing_row)
		url_dict = dict(zip(self.keys_list, value_list))
		if not url_dict.get('url_trimmed'):
			return url_dict
		return None

	def trim_url(self, url_dict):

		url = url_dict.get('url')

		try:
			if 'www.' in url:
				url = re.match('.*?www.(.*)/', url).group(1)
			else:
				url = re.match('.*?//(.*)/', url).group(1)
			url_dict['url_trimmed'] = url
			return url_dict
		except:
			logger.error('正则匹配失败 ' + url)
			return None

	def mark_url(self, url_dict):

		sheet = self.workbook_write.get_sheet(self.sheet_order)
		sheet.write(self.processing_row, self.url_trimmed_coloumn, url_dict.get('url_trimmed'))
		self.workbook_write.save(self.workbook_title)
		logger.info(str(self.processing_row)+ ' ' + url_dict.get('url_trimmed'))

	def run(self):

		row_number = self.sheet_data.nrows - 1

		for row in range(row_number):
			url_dict = self.fetch_url_undone()
			if url_dict:
				url_dict = self.trim_url(url_dict)
				if url_dict:
					self.mark_url(url_dict)

if __name__ == '__main__':

	trim_urls = TrimUrls()
	trim_urls.run()

