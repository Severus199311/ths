#coding: utf-8
#微信经常被封，用来拉出被封帐号专注的种子

import requests, xlwt, xlrd
from xlutils.copy import copy
from logger import logger
from urllib.parse import urlencode

class SortSeeds():

	def __init__(self):
		
		self.base_url = 'http://flashcms.10jqka.com.cn/seed/seed/searchSeed/?'
		self.accounts_blacklisted_list = ['wxid_rqf2jl8qe2je22','wxid_it55uv5507s122','wxid_52xxcwf9qasb22','wxid_8lw80u8r6asr22','wxid_agzvm4fvui9n22','wxid_q8bo3ba5b4dr22','wxid_j2loegxbjba322','wxid_gi18dps0jn6w22','wxid_chv90te5nf6222','wxid_dg79hhnshq4d22','wxid_3fk9a7y3vk7822','wxid_fhyrfjo2s6n522','wxid_amn25xsheiea22','wxid_xf22ddtqsvjq22','wxid_wjkpo6chcxs422','wxid_y92tijtllocp12','wxid_rspzklg0w0qz12','wxid_1fpwg8gmu9p22','wxid_ty1nu2avjwhb22','wxid_fpdvvxwbvmp912','wxid_pzjxdtpb41w212','wxid_a24sypdkt8t112','wxid_orkd448qbinu12','wxid_u37bgw88kzry22','wxid_1zcfgkb8y7w222','wxid_yhu4lw27j3g622','wxid_fb7nitaowk9h22','wxid_pjxpczmxwlfg22','wxid_ulpsv4eb49t422','wxid_nnyvoq8taqjn22','wxid_t2r6vmta18p312']

		#搜索条件：
		self.params = {
			'keywords': '', #关键词
			'platform': 'wechat', #平台，分为：web,wechat,weibo,app,rss
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
		self.total_pages = 466 #总的页面数。要去前段页面找。可以根据pageSize作出调整。
		self.start_page = 1

		self.workbook_title = 'wechat_seeds_sorted.xls'
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
			seed_dict['scheduler_name'] = data.get('schedulerName')
			seed_dict['disabled'] = data.get('disabled')
			seed_dict['user_name'] = data.get('userName')
			seeds_list.append(seed_dict)
		return seeds_list

	#判断公众号是否只有被封的帐号关注
	def if_account_blacklisted(self, seed_dict):

		if not seed_dict.get('user_name'):
			return True

		else:
			user_name_list = seed_dict.get('user_name').split(',')
			for user_name in user_name_list:
				if user_name not in self.accounts_blacklisted_list:
					return False
			return True

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

		for page in range(self.start_page, self.total_pages+1):

			this_params = self.params
			this_params['page'] = str(page)
			logger.info('准备抓取页面：' + str(page))
			url = self.build_url(this_params)
			raw_data = self.get_page(url)
			if raw_data:
				seeds_list = self.parse_page(raw_data)
				for seed_dict in seeds_list:
					if self.if_account_blacklisted(seed_dict):
						logger.info(seed_dict)
						#self.write_into_excel(seed_dict)

if __name__ == '__main__':
	sort_seeds = SortSeeds()
	sort_seeds.run()