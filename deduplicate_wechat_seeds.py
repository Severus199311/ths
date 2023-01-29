#根据url测试种子是否存在

import xlwt, xlrd, re, requests
from xlutils.copy import copy
from logger import logger


class DedeplicateSeedsByURL():

	def __init__(self):

		self.headers_seeds = {
			'Cookie': 'PHPSESSID=tft1f9dlha2uclqmm4oi42q9d4'
		}
		self.base_url = 'http://flashcms.10jqka.com.cn/seed/seed/searchSeed/?platform=wechat&disabled=&keywords={}&central=1&page=1&sourceid=&pageSize=100'

		self.workbook_title = 'wechat_seeds_to_be_deduplicated.xls' #文件
		self.sheet_name = 'Sheet1' #表名
		self.sheet_order = 0 #表序号

		self.if_successful_column_order = 1 #if_successful在第2列，所以是1
		self.if_successful_marker = 1
		self.seed_id_column_order = 3
		self.status_column_order = 4

		self.workbook_read = xlrd.open_workbook(self.workbook_title) #打开excel文件，用于读取数据
		self.sheet_data = self.workbook_read.sheet_by_name(self.sheet_name) #读取表数据
		self.keys_list = self.sheet_data.row_values(0) #获取列名，格式为list
		self.workbook_write = copy(self.workbook_read) #复制excel文件，用于写入
		self.processing_row = 0 #初始行序号为0

	def fetch_account_dict(self):

		self.processing_row += 1
		value_list = self.sheet_data.row_values(self.processing_row)
		account_dict = dict(zip(self.keys_list, value_list))
		if not account_dict.get('if_successful'):
			return account_dict
		return None

	def get_seed_urls(self, keywords):

		seed_url = self.base_url.format(keywords)
		account_dict = {'account_name': keywords}

		try:
			response = requests.get(seed_url, headers=self.headers_seeds)
		except Exception:
			logger.error('请求失败')
			response = None

		if response:
			data_list = response.json().get('result').get('list')
			if data_list:
				for data in data_list:
					if data.get('siteName') == keywords:
						account_dict['seed_id'] = data.get('id')
						account_dict['status'] = data.get('status')
		return account_dict

	def mark_account_deduplicated(self, account_dict):

		sheet = self.workbook_write.get_sheet(self.sheet_order)
		sheet.write(self.processing_row, self.if_successful_column_order, self.if_successful_marker)
		if account_dict.get('seed_id'):
			sheet.write(self.processing_row, self.seed_id_column_order, str(account_dict.get('seed_id')))
			sheet.write(self.processing_row, self.status_column_order, str(account_dict.get('status')))
		self.workbook_write.save(self.workbook_title)
		logger.info(account_dict)

	def run(self):

		row_number = self.sheet_data.nrows - 1
		for row in range(row_number):
			account_dict = self.fetch_account_dict()
			if account_dict:
				keywords = account_dict.get('account_name')
				account_dict = self.get_seed_urls(keywords)
				self.mark_account_deduplicated(account_dict)

if __name__ == '__main__':
	deduplicate_seeds_by_url = DedeplicateSeedsByURL()
	deduplicate_seeds_by_url.run()