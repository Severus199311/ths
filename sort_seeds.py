#coding: utf-8
#根据条件获取种子

import requests, xlwt, xlrd
from xlutils.copy import copy
from logger import logger
from urllib.parse import urlencode

class SortSeeds():

	def __init__(self):
		
		self.base_url = 'http://flashcms.10jqka.com.cn/seed/seed/searchSeed/?'

		#搜索条件：
		self.params = {
			'keywords': '基金', #关键词
			'platform': 'web', #平台，分为：web,wechat,weibo,app,rss
			'star': '', #星级。1，2，3，4，5
			'group': '', #分组序号可能会变，每次都要去前端页面确认
			'disabled': '', #状态。0，1，2，3，也可以没有
			'central': '1', #不知道是什么
			'page': '', #页面。可以和pagesize配和
			'sourceid': '',
			'pageSize': '100' #每页的种子数。可以和page配合，比如：共有500条种子，可设置为page 1至10、pageSize 50，也可设置为page 1至25、pageSize 20。pageSize此前最大可设置为100
		}

		self.headers = {
			'Cookie': 'PHPSESSID=tft1f9dlha2uclqmm4oi42q9d4'
		}
		self.total_pages = 9 #总的页面数。要去前段页面找。可以根据pageSize作出调整。

		self.workbook_title = 'seeds_sorted.xls'
		self.sheet_name = 'Sheet1' #待录入种子所在的表名
		self.sheet_order = 0 #待录入种子所在的表序号

	def build_url(self, params):
		return self.base_url + urlencode(self.params)

	def get_page(self, url):
		try:
			response = requests.get(url, headers=self.headers)
		except Exception as e:
			logger.error('请求失败——' + e.args[0])
			return None

		if response.status_code == 200:
			return response.json()
		else:
			logger.error('无效响应吗——' + str(response.status_code))
			return None

	def parse_page(self, raw_data):
		seeds_list = []
		data_list = raw_data.get('result').get('list')
		for data in data_list:
			seed_dict = {}
			#seed_dict['order'] = ''
			#seed_dict['if_successful'] = ''
			seed_dict['seed_id'] = data.get('id')
			seed_dict['seed_title'] = data.get('comments')
			seed_dict['url'] = data.get('url')
			seed_dict['star'] = data.get('star')
			seed_dict['group'] = data.get('group')
			seed_dict['tags'] = data.get('tag')
			seed_dict['platform'] = data.get('platform')
			seed_dict['schedulerName'] = data.get('schedulerName')
			seed_dict['disabled'] = data.get('disabled')
			seeds_list.append(seed_dict)
		return seeds_list

	def write_into_excel(self, seed_dict):
		workbook_read = xlrd.open_workbook(self.workbook_title) #打开excel文件，用于读取数据
		workbook_write = copy(workbook_read) #复制excel文件，用于写入
		sheet_data = workbook_read.sheet_by_name(self.sheet_name) #读取表数据
		row_number = sheet_data.nrows
		sheet_write = workbook_write.get_sheet(self.sheet_order) #用于写入的表格
		data_list = list(seed_dict.values())
		for data in data_list:
			sheet_write.write(row_number, data_list.index(data), data)
		workbook_write.save(self.workbook_title)

	def run(self):
		for page in range(1, self.total_pages+1):
			this_params = self.params
			this_params['page'] = str(page)
			logger.info('准备抓取页面：' + str(page))
			url = self.build_url(this_params)
			raw_data = self.get_page(url)
			if raw_data:
				seeds_list = self.parse_page(raw_data)
				for seed_dict in seeds_list:
					self.write_into_excel(seed_dict)

	def run_2(self):

		keywords_list = ['Ecns','IT之家','Yicai Global','ZAKER新闻','北京时间','财新','参考消息','畅说108','第一财经','第一财经周刊','动漫之家','观察者','国务院','海客新闻','虎扑','虎嗅','华尔街见闻','界面新闻','蓝鲸财经','礼堂家','荔枝新闻','南方Plus','南方周末','上观','生物谷','世界浙商网','四川新闻','搜狐','钛媒体','网易','文汇','问津','携程','新湖南','新京报','新浪体育','新民','游戏时光','证券时报','知乎日报','中国新闻网']
		
		keyword = keywords_list.pop()
		page = 1
		while True:
			this_params = self.params
			this_params['keywords'] = keyword
			this_params['page'] = str(page)
			logger.info('准备抓取：' + keyword + ' ' + str(page))
			url = self.build_url(this_params)
			raw_data = self.get_page(url)
			if raw_data:
				seeds_list = self.parse_page(raw_data)
				if len(seeds_list) == 0:
					if len(keywords_list) == 0:
						break
					else:
						keyword = keywords_list.pop()
						page = 1
				else:
					for seed_dict in seeds_list:
						self.write_into_excel(seed_dict)
					page += 1

if __name__ == '__main__':
	sort_seeds = SortSeeds()
	sort_seeds.run()