#获取种子库里的微信种子，用来给微信扩源去重

import sys
import time
import xlwt, xlrd
from xlutils.copy import copy
import requests
from logger import logger

class SortWechatSeeds():

	def __init__(self):
		self.workbook_title = 'wechat_seeds.xls'
		self.sheet_name = 'Sheet1' #待录入种子所在的表名
		self.sheet_order = 0 #待录入种子所在的表序号
		self.columns_dict = {'account_name': 1, 'account_id': 2,'seed_id': 3, 'groups': 4, 'tags': 5} #key是字段，value是字段所在列的序号。写入的时候需要
		self.headers = {'Cookie': 'PHPSESSID=2i9m1upvem6vj25og0apb02em0'}
		self.page_size = 100 #每页获取种子的个数
		self.start_page_number = 414 #首次运行程序时，从第一页开始，此后每次从上一次的结束页开始
		#self.base_url = 'http://flashcms.10jqka.com.cn/seed/seed/searchSeed/?platform=wechat&disabled=&keywords=&central=1&page={}&sourceid=&pageSize={}'
		#self.base_url = 'http://flashcms.10jqka.com.cn/seed/seed/searchSeed/?platform=wechat&disabled=2&keywords=&central=1&page={}&sourceid=&pageSize={}'
		#self.base_url = 'http://flashcms.10jqka.com.cn/seed/seed/searchSeed/?platform=wechat&disabled=0&keywords=&central=1&page={}&sourceid=&pageSize={}'
		self.base_url = 'http://flashcms.10jqka.com.cn/seed/seed/searchSeed/?platform=wechat&star=5&disabled=0&keywords=&central=1&page=1&sourceid=&pageSize={}'

	def get_seed_page(self,page_number):

		url = self.base_url.format(page_number, self.page_size)
		response = requests.get(url ,headers=self.headers)
		if response.status_code == 200:
			return response.json()
		logger.info('请求失败，响应码：' + str(response.status_code))
		return None

	def parse_seed_page(self, seed_page):
		seeds_list = []
		data_list = seed_page.get('result').get('list')
		if not data_list:
			return None
		for data in data_list:
			seed_dict = {}
			seed_dict['account_name'] = data.get('siteName')
			seed_dict['account_id'] = data.get('url')
			seed_dict['seed_id'] = data.get('id')
			if data.get('group'):
				seed_dict['groups'] = ','.join(data.get('group'))
			if data.get('tag'):
				seed_dict['tags'] = ','.join(data.get('tag'))
			seeds_list.append(seed_dict)
		return seeds_list

	def check_key_words(self, seed_dict):

		for key_word in ['大金融', '法制', '搞笑', '分析师', '新闻', '历史', '旅游', '情感', '育儿', '债', '政', '自然', '职', '游戏', '协会', '期货', '时尚', '美食', '风水', '公司','部委','政府','股','交易所','板','中央','地方','竞品','公告','文','境外','视频号','海外','基金']:
			if seed_dict.get('tags') and key_word in seed_dict.get('tags'):
				return True
		return False

	def write_into_excel(self, seed_dict):

		workbook_read = xlrd.open_workbook(self.workbook_title) #打开excel文件，用于读取数据
		workbook_write = copy(workbook_read) #复制excel文件，用于写入
		sheet_data = workbook_read.sheet_by_name(self.sheet_name) #读取表数据
		row_number = sheet_data.nrows
		sheet_write = workbook_write.get_sheet(self.sheet_order) #用于写入的表格
		for key, value in self.columns_dict.items():
			#if not self.check_key_words(seed_dict):
			#	sheet_write.write(row_number, value, seed_dict.get(key))
			sheet_write.write(row_number, value, seed_dict.get(key))
		workbook_write.save(self.workbook_title)

	def run(self):

		while True:
			seed_page = self.get_seed_page(self.start_page_number)
			if seed_page:
				seeds_list = self.parse_seed_page(seed_page)
				if not seeds_list:
					logger.info('页码：' + str(self.start_page_number) + ' 种子列表为空')
				else:
					for seed_dict in seeds_list:
						self.write_into_excel(seed_dict)
					logger.info('页码：' + str(self.start_page_number) + ' ' + '获取种子数：' + str(len(seeds_list)))
				self.start_page_number += 1

if __name__ == '__main__':
	sort_wechat_seeds = SortWechatSeeds()
	sort_wechat_seeds.run()