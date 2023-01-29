#将excel中的站点名与媒体库的媒体名称、媒体别名对比，去重

import xlwt, xlrd
from xlutils.copy import copy
import time, re, json, requests, sys
from logger import logger

class CheckMedia():

	def __init__(self):

		self.workbook_title = 'media_to_be_checked.xls' #待录入媒体所在的文件
		self.sheet_name = 'Sheet1' #待录入媒体所在的表名
		self.sheet_order = 0 #待录入媒体所在的表序号

		self.search_existing_media_url = 'http://flashcms.10jqka.com.cn/entry/media/ajax'
		self.search_existing_seeds_base_url = 'http://flashcms.10jqka.com.cn/seed/seed/searchSeed/?platform=&disabled=&keywords={}&central=1&page=1&sourceid=&pageSize=20'
		self.search_existing_media_headers = {
			'Cookie': 'PHPSESSID=8sa2mekqcj9ml8g9fe4agmuju7; __gads=ID=697ce974af639e2a:T=1626770434:S=ALNI_MZjlS6_ZYxo9W_Vk4hFbsrPoF19EA'
		}
		self.search_existing_seeds_headers = {
			'Cookie': 'PHPSESSID=8sa2mekqcj9ml8g9fe4agmuju7; __gads=ID=697ce974af639e2a:T=1626770434:S=ALNI_MZjlS6_ZYxo9W_Vk4hFbsrPoF19EA'
		}
		#以上二项供search_existing_media方法使用

		self.workbook_read = xlrd.open_workbook(self.workbook_title) #打开excel文件，用于读取数据
		self.sheet_data = self.workbook_read.sheet_by_name(self.sheet_name) #读取表数据
		self.keys_list = self.sheet_data.row_values(0) #获取列名，格式为list
		self.workbook_write = copy(self.workbook_read) #复制excel文件，用于写入
		self.processing_row = 0 #初始行序号为0

	#从excel表格中获取待录入种子信息
	def fetch_media_unchecked(self):
		self.processing_row += 1
		value_list = self.sheet_data.row_values(self.processing_row)
		media_dict = dict(zip(self.keys_list, value_list))
		if not media_dict.get('if_successful'):
			return media_dict
		return None

	def search_existing_media_and_seeds(self, media_dict):

		media_name = media_dict.get('stock_name')

		d_1 = {"page":"1","site":"","name":"","alias":"","tag":"","type":"","weight":"","recent":"","iscopyright":"","iscopystart":None,"iscopyend":None}
		d_2 = {"page":"1","site":"","name":"","alias":"","tag":"","type":"","weight":"","recent":"","iscopyright":"","iscopystart":None,"iscopyend":None}
		d_1['name'] = media_name
		d_2['alias'] = media_name
		data_1 = {
			'opt': 'search',
			'd': json.dumps(d_1)
		}
		data_2 = {
			'opt': 'search',
			'd': json.dumps(d_2)
		}

		response_a = requests.post(self.search_existing_media_url, headers=self.search_existing_media_headers, data=data_1)
		if response_a.status_code == 200:
			media_list_a = response_a.json().get('data').get('data')
			if media_list_a:
				media_names_a = media_domains_a = media_aliases_a = ''
				for media_a in media_list_a:
					media_names_a += media_a.get('name')
					media_names_a += ','
					media_domains_a += media_a.get('linkurl')
					media_domains_a += ','
					media_aliases_a += media_a.get('alias')
					media_domains_a += ','
				media_dict['media_names_a'] = media_names_a
				media_dict['media_domains_a'] = media_domains_a
				media_dict['media_aliases_a'] = media_aliases_a
		else:
			logger.error('请求媒体库失败，请检查cookie')
			sys.exit()

		response_b = requests.post(self.search_existing_media_url, headers=self.search_existing_media_headers, data=data_2)
		if response_b.status_code == 200:
			media_list_b = response_b.json().get('data').get('data')
			if media_list_b:
				media_names_b = media_domains_b = media_aliases_b = ''
				for media_b in media_list_b:
					media_names_b += media_b.get('name')
					media_names_b += ','
					media_domains_b += media_b.get('linkurl')
					media_domains_b += ','
					media_aliases_b += media_b.get('alias')
					media_aliases_b += ','
				media_dict['media_names_b'] = media_names_b
				media_dict['media_domains_b'] = media_domains_b
				media_dict['media_aliases_b'] = media_aliases_b
		else:
			logger.error('请求媒体库失败，请检查cookie')
			sys.exit()

		search_existing_seeds_url = self.search_existing_seeds_base_url.format(media_name)
		response_c = requests.get(search_existing_seeds_url, headers=self.search_existing_seeds_headers)
		if response_c.status_code == 200:
			seeds_list = response_c.json().get('result').get('list')
			if seeds_list:
				seed_names = seed_urls = ''
				for seed in seeds_list:
					seed_names += seed.get('comments')
					seed_names += ','
					seed_urls += seed.get('landingPage')
					seed_urls += ','
				media_dict['seed_names'] = seed_names
				media_dict['seed_urls'] = seed_urls
		else:
			logger.error('请求种子库失败，请检查cookie')
			sys.exit()

		return media_dict

	def mark_media_done(self, media_dict):
		sheet = self.workbook_write.get_sheet(self.sheet_order)
		sheet.write(self.processing_row, 1, 1)
		sheet.write(self.processing_row, 4, media_dict['media_names_a'])
		sheet.write(self.processing_row, 5, media_dict['media_domains_a'])
		sheet.write(self.processing_row, 6, media_dict['media_aliases_a'])
		sheet.write(self.processing_row, 7, media_dict['media_names_b'])
		sheet.write(self.processing_row, 8, media_dict['media_domains_b'])
		sheet.write(self.processing_row, 9, media_dict['media_aliases_b'])
		sheet.write(self.processing_row, 10, media_dict['seed_names'])
		sheet.write(self.processing_row, 11, media_dict['seed_urls'])
		self.workbook_write.save(self.workbook_title)
		logger.info(str(self.processing_row)+ ' ' + media_dict['media_names_a'] + ' ' + media_dict['media_names_b'] + ' ' + media_dict['seed_names'])

	def run(self):

		row_number = self.sheet_data.nrows - 1
		for row in range(row_number):
			media_dict = self.fetch_media_unchecked()
			if media_dict: #已经标注过的，不会再进行下面代码
				media_dict = self.search_existing_media_and_seeds(media_dict)
				self.mark_media_done(media_dict)

if __name__ == '__main__':
	check_media = CheckMedia()
	check_media.run()