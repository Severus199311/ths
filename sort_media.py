#整理出媒体信息

import xlwt, xlrd
from xlutils.copy import copy
import time, re, json, requests, sys
from logger import logger

class SortMedia():

	def __init__(self):

		self.workbook_title = 'media_sorted.xls' #待录入媒体所在的文件
		self.sheet_name = 'Sheet1' #待录入媒体所在的表名
		self.sheet_order = 0 #待录入媒体所在的表序号
		self.start_page = 49
		self.total_pages = 1037

		self.search_media_url = 'http://flashcms.10jqka.com.cn/entry/media/ajax'
		self.search_media_headers = {
			'Cookie': 'PHPSESSID=shoqql0hgm4ecknssn04sn7a12'
			}
		self.columns_dict = {'name': 0, 'alias' : 1, 'url': 2, 'site': 3, 'tag': 4, 'type': 5, 'weight': 6, 'copyright': 7}

	def search_media(self, page):

		d = {"page":"","site":"","name":"","alias":"","tag":"","type":"","weight":"","recent":"","iscopyright":"","iscopystart":None,"iscopyend":None}
		d['page'] = str(page)

		data = {
			'opt': 'search',
			'd': json.dumps(d)
		}

		response = requests.post(self.search_media_url, headers=self.search_media_headers, data=data)

		if response.status_code == 200:
			media_list = response.json().get('data').get('data')

			if media_list:
				media_dict_list = []
				for media in media_list:
						media_dict = {}
						media_dict['name'] = media.get('name')
						media_dict['alias'] = media.get('alias')
						media_dict['url'] = media.get('linkurl').replace('|', ',')
						media_dict['site'] = media.get('site')
						media_dict['tag'] = media.get('tag')
						media_dict['type'] = media.get('type')
						media_dict['weight'] = media.get('weight')
						media_dict['copyright'] = media.get('iscopyright')
						media_dict_list.append(media_dict)
				return media_dict_list

			else:
				logger.error('该页列表为空，请检查搜索关键词')
				sys.exit()

		else:
			logger.error('请求媒体库失败，请检查cookie')
			sys.exit()

	def write_into_excel(self, media_dict):

		workbook_read = xlrd.open_workbook(self.workbook_title) #打开excel文件，用于读取数据
		workbook_write = copy(workbook_read) #复制excel文件，用于写入
		sheet_data = workbook_read.sheet_by_name(self.sheet_name) #读取表数据
		row_number = sheet_data.nrows
		sheet_write = workbook_write.get_sheet(self.sheet_order) #用于写入的表格
		for key, value in self.columns_dict.items():
			sheet_write.write(row_number, value, media_dict.get(key))
		workbook_write.save(self.workbook_title)

	def run(self):

		for page in range(self.start_page, self.total_pages + 1):

			print(page)

			media_dict_list = self.search_media(str(page))

			for media_dict in media_dict_list:

				self.write_into_excel(media_dict)

if __name__ == '__main__':

	sort_media = SortMedia()
	sort_media.run()
