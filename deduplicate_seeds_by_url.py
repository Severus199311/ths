#根据url测试种子是否存在

import xlwt, xlrd, re, requests
from xlutils.copy import copy
from logger import logger


class DedeplicateSeedsByURL():

	def __init__(self):

		self.headers_seeds = {
			'Cookie': 'PHPSESSID=hhik9dpf220kcqt0qru0c75vo4'
		}
		self.base_url = 'http://flashcms.10jqka.com.cn/seed/seed/searchSeed/?platform=web&disabled=&keywords={}&central=1&page=1&sourceid=&pageSize=100'

		self.workbook_title = 'seeds_to_be_deduplicated.xls' #文件
		self.sheet_name = 'Sheet1' #表名
		self.sheet_order = 0 #表序号

		self.if_successful_column_order = 1 #if_successful在第2列，所以是1
		self.if_successful_marker = 1
		self.if_already_exist_column_order = 5

		self.workbook_read = xlrd.open_workbook(self.workbook_title) #打开excel文件，用于读取数据
		self.sheet_data = self.workbook_read.sheet_by_name(self.sheet_name) #读取表数据
		self.keys_list = self.sheet_data.row_values(0) #获取列名，格式为list
		self.workbook_write = copy(self.workbook_read) #复制excel文件，用于写入
		self.processing_row = 0 #初始行序号为0

	def fetch_url_dict(self):

		self.processing_row += 1
		value_list = self.sheet_data.row_values(self.processing_row)
		url_dict = dict(zip(self.keys_list, value_list))
		if not url_dict.get('if_successful'):
			return url_dict
		return None

	def trim_url(self, url):

		try:
			url = re.match('http.*?//(.*)', url).group(1)
			if re.match('^www\..*', url):
				url = re.match('^www\.(.*)', url).group(1)
			if re.match('.*#$', url):
				url = re.match('(.*)#$', url).group(1)
			if re.match('.*\.shtml$', url):
				url = re.match('(.*)\.shtml$', url).group(1)
			if re.match('.*\.html$', url):
				url = re.match('(.*)\.html$', url).group(1)
			if re.match('.*\.htm$', url):
				url = re.match('(.*)\.htm$', url).group(1)
			if re.match('.*/$', url):
				url = re.match('(.*)/$', url).group(1)
			if re.match('.*index$', url):
				url = re.match('(.*)index$', url).group(1)
			return url
		except:
			logger.error('正则匹配失败 ' + url)
			return None

	def get_seed_urls(self, keywords):

		seed_url = self.base_url.format(keywords)

		try:
			response = requests.get(seed_url, headers=self.headers_seeds)
		except Exception:
			logger.error('请求失败')
			response = None

		if response:
			data_list = response.json().get('result').get('list')
			if data_list:
				urls_list = []
				for data in data_list:
					if self.trim_url(data.get('url')) == keywords:
						#urls_list.append(data.get('url'))
						urls_list.append(str(data.get('id')))			
				return ', '.join(urls_list)
		return None

	def mark_url_deduplicated(self, url_dict):

		sheet = self.workbook_write.get_sheet(self.sheet_order)
		sheet.write(self.processing_row, self.if_already_exist_column_order, url_dict.get('if_already_exist'))
		sheet.write(self.processing_row, self.if_successful_column_order, self.if_successful_marker)
		self.workbook_write.save(self.workbook_title)
		if url_dict.get('if_already_exist'):
			logger.info(str(self.processing_row) + ' ' + url_dict.get('url') + ' 存在重复: ' + url_dict.get('if_already_exist'))
		else:
			logger.info(str(self.processing_row)+ ' ' + url_dict.get('url'))


	def run(self):

		row_number = self.sheet_data.nrows - 1
		for row in range(row_number):
			url_dict = self.fetch_url_dict()
			if url_dict:
				keywords = self.trim_url(url_dict.get('url'))
				if keywords:
					url_dict['if_already_exist'] = self.get_seed_urls(keywords)
					self.mark_url_deduplicated(url_dict)


if __name__ == '__main__':
	deduplicate_seeds_by_url = DedeplicateSeedsByURL()
	deduplicate_seeds_by_url.run()