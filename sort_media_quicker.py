#整理出媒体信息.由于先把数据全部获取，在全部写入，在保存excel，会比sort_media快很多

import xlwt, xlrd
from xlutils.copy import copy
import time, re, json, requests, sys
from logger import logger

class SortMedia():

	def __init__(self):

		self.start_page = 1
		self.total_pages = 1039

		self.search_media_url = 'http://flashcms.10jqka.com.cn/entry/media/ajax'
		self.search_media_headers = {
			'Cookie': 'PHPSESSID=2i9m1upvem6vj25og0apb02em0'
			}
		self.columns_dict = {'name': 0, 'alias' : 1, 'url': 2, 'site': 3, 'tag': 4, 'type': 5, 'weight': 6, 'copyright': 7}

		self.media_dict_list = []

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
						self.media_dict_list.append(media_dict)

			else:
				logger.error('该页列表为空，请检查搜索关键词')
				sys.exit()

		else:
			logger.error('请求媒体库失败，请检查cookie')
			sys.exit()

	def write_into_excel(self):

		workbook = xlwt.Workbook(encoding='utf-8')
		worksheet = workbook.add_sheet('Sheet1')
		for i in range(len(self.media_dict_list)):
			media_dict = self.media_dict_list[i]
			worksheet.write(i, 0, media_dict.get('name'))
			worksheet.write(i, 1, media_dict.get('alias'))
			worksheet.write(i, 2, media_dict.get('url'))
			worksheet.write(i, 3, media_dict.get('site'))
			worksheet.write(i, 4, media_dict.get('tag'))
			worksheet.write(i, 5, media_dict.get('type'))
			worksheet.write(i, 6, media_dict.get('weight'))
			worksheet.write(i, 7, media_dict.get('copyright'))
		workbook.save('media_sorted.xls')

	def run(self):

		for page in range(self.start_page, self.total_pages + 1):
			logger.info('正在抓取页面' + str(page))
			self.search_media(str(page))

		logger.info('正在写入excel')
		self.write_into_excel()

if __name__ == '__main__':

	sort_media = SortMedia()
	sort_media.run()
