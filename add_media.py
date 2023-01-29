#coding=utf-8
#判断媒体是否已经存在，不存在就配媒体

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC 
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
import xlwt, xlrd
from xlutils.copy import copy
import time, re, json, requests
from logger import logger

class AddMedia():

	def __init__(self):

		self.workbook_title = 'media.xls' #待录入媒体所在的文件
		self.sheet_name = 'Sheet1' #待录入媒体所在的表名
		self.sheet_order = 0 #待录入媒体所在的表序号
		#以上三项为最关键配置，忘记修改的话，媒体录入将出错，且无法修复

		self.if_successful_column_order = 1 #if_successful在第2列，所以是1
		self.new_media_marker = 1 #成功，则标记为1
		self.media_name_exists_marker = 2 #媒体名存在，则标记为2。但是，如果被测媒体名仅仅是已有媒体名的一部分，也会被认为存在，例如搜索‘自然资源’，会出现‘自然资源部’，因此也会被认为存在。 下面别名也一样
		self.media_alias_exists_marker =3 #媒体别名存在，则标记为3
		#以上四项也是重要配置，忘记修改的话，excel记录将出错，但可以修复

		self.start_url = 'http://flashcms.10jqka.com.cn/default/index/index'
		self.start_url_2 = 'http://flashcms.10jqka.com.cn/entry/media/edit/'
		self.wait_time = 30 #selenium等待时长
		self.sleep_time = 0.3

		self.search_existing_media_url = 'http://flashcms.10jqka.com.cn/entry/media/ajax'
		self.search_existing_media_headers = {
			'Cookie': 'PHPSESSID=p64csf8cf53d2utlsa18101so0'
			}
		#以上二项供search_existing_media方法使用

		self.workbook_read = xlrd.open_workbook(self.workbook_title) #打开excel文件，用于读取数据
		self.sheet_data = self.workbook_read.sheet_by_name(self.sheet_name) #读取表数据
		self.keys_list = self.sheet_data.row_values(0) #获取列名，格式为list
		self.workbook_write = copy(self.workbook_read) #复制excel文件，用于写入
		self.processing_row = 0 #初始行序号为0

		self.browser = webdriver.Chrome() 
		self.wait = WebDriverWait(self.browser, self.wait_time)

	#打开媒体录入页面
	def open_page(self):
		self.browser.get(self.start_url)
		self.wait.until(EC.alert_is_present())
		self.browser.switch_to.alert.accept()
		self.wait.until(EC.alert_is_present())
		self.browser.switch_to.alert.accept()
		#self.wait.until(EC.presence_of_element_located((By.ID, '_M7')))
		self.browser.get(self.start_url_2)
		self.browser.maximize_window()

	#从excel表格中获取待录入种子信息
	def fetch_media_undone(self):
		self.processing_row += 1
		value_list = self.sheet_data.row_values(self.processing_row)
		media_dict = dict(zip(self.keys_list, value_list))
		if not media_dict.get('if_successful'):
			#return media_dict
			#if media_dict.get('media_tag') and media_dict.get('media_type') and media_dict.get('media_weight') and media_dict.get('copyright'): 
			if media_dict.get('media_use') and media_dict.get('media_type') and media_dict.get('media_weight') and media_dict.get('copyright'): 
				return media_dict
		return None

	#判断媒体名和媒体别名是否已经存在
	def search_existing_media(self, media_name):

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

		response = requests.post(self.search_existing_media_url, headers=self.search_existing_media_headers, data=data_1)
		if response.status_code == 200:
			media_list = response.json().get('data').get('data')
			if media_list:
				for media in media_list:
					if media.get('name') == media_name:
						return self.media_name_exists_marker
		else:
			logger.error('请求媒体库失败，请检查cookie')

		response = requests.post(self.search_existing_media_url, headers=self.search_existing_media_headers, data=data_2)
		if response.status_code == 200:
			media_list = response.json().get('data').get('data')
			if media_list:
				for media in media_list:
					if media.get('alias') == media_name:
						return self.media_alias_exists_marker
		else:
			logger.error('请求媒体库失败，请检查cookie')

		return self.new_media_marker

	#将新处理好的媒体标记为1
	def mark_media_done(self, media_dict):
		
		sheet = self.workbook_write.get_sheet(self.sheet_order)
		sheet.write(self.processing_row, self.if_successful_column_order, media_dict.get('marker')) #该标记其实有1，2，3，分别对应录入成功、媒体名存在、媒体别名存在
		self.workbook_write.save(self.workbook_title)
		logger.info(str(self.processing_row)+ ' ' + media_dict.get('media_name') + ' 标记为' + str(media_dict.get('marker')))

	def wait_and_get_element(self, scenario, css_path):
		if scenario == 'located':
			return self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, css_path)))
		if scenario == 'clickable':
			return self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, css_path)))
			

	#录入媒体
	def add_media(self, media_dict):

		#标记媒体站点
		if media_dict.get('media_use'):
			media_use = self.wait_and_get_element('located', '#site')
			media_use.clear()
			media_use.send_keys(media_dict.get('media_use'))
			
		media_title = self.wait_and_get_element('located', '#name')
		media_title.clear()
		media_title.send_keys(media_dict.get('media_name'))

		#"""
		media_tag = self.wait_and_get_element('clickable', '#tag')
		media_tag.click()
		media_tags_list = self.wait_and_get_element('clickable', '#tag option')
		media_tags_list = self.browser.find_elements_by_css_selector('#tag option')

		
		media_tag_marked = False
		for each_tag in media_tags_list:
			if each_tag.text == media_dict.get('media_tag'):
				each_tag.click()
				media_tag_marked = True
				break
		if not media_tag_marked:
			return None
		#"""

		media_type = self.wait_and_get_element('clickable', '#type')
		media_type.click()
		media_types_list = self.wait_and_get_element('clickable', '#type option')
		media_types_list = self.browser.find_elements_by_css_selector('#type option')

		media_type_marked  = False
		for each_type in media_types_list:
			if each_type.text == media_dict.get('media_type'):
				each_type.click()
				media_type_marked = True
				break
		if not media_type_marked: #存在可能，excel中的type与媒体库任何type都对不上，因此需要这个判断
			return None

		weights = self.wait_and_get_element('clickable', '#weight')
		weights.click()
		weights_list = self.wait_and_get_element('clickable', '#weight option')
		weights_list = self.browser.find_elements_by_css_selector('#weight option')

		weight_marked = False
		for each_weight in weights_list:
			if each_weight.text == media_dict.get('media_weight'):
				each_weight.click()
				weight_marked = True
				break
		if not weight_marked: #存在可能，excel中的weight与媒体库任何weight都对不上，因此需要这个判断
			return None

		copyrights = self.wait_and_get_element('clickable', '#iscopyright')
		copyrights.click()
		copyrights_list = self.wait_and_get_element('clickable', '#iscopyright option')
		copyrights_list = self.browser.find_elements_by_css_selector('#iscopyright option')

		copyright_marked = False
		for each_copyright in copyrights_list:
			if each_copyright.text == media_dict.get('copyright'):
				each_copyright.click()
				copyright_marked = True
				break
		if not copyright_marked: #存在可能，excel中的copyright与媒体库任何copyright都对不上，因此需要这个判断
			return None

		account_intro = self.wait_and_get_element('located', '#intro')
		account_intro.clear()
		if media_dict.get('account_intro'):
			account_intro.send_keys(media_dict.get('account_owner') + '；' + media_dict.get('account_intro'))

		#标记别名
		alias = self.wait_and_get_element('located', '#alias')
		alias.clear()
		alias_list = media_dict.get('alias').split(',')
		for each in alias_list:
			alias.send_keys(each)
			alias.send_keys(Keys.ENTER)


		#标记稿源
		source = self.wait_and_get_element('located', '#linkurl')
		source.clear()
		source_list = media_dict.get('source').split(',')
		for each in source_list:
			source.send_keys(each)
			source.send_keys(Keys.ENTER)

		submit = self.wait_and_get_element('clickable', '#btn-fd')
		submit.click()

		self.wait.until(EC.alert_is_present())
		self.browser.switch_to.alert.dismiss()

		return media_dict

	def run(self):
		
		self.open_page()
		row_number = self.sheet_data.nrows - 1
		for row in range(row_number):
			media_dict = self.fetch_media_undone()
			if media_dict: #已经标注过的，不会再进行下面代码
				media_dict['marker'] = self.search_existing_media(media_dict.get('media_name'))
				if media_dict['marker'] == self.new_media_marker: #如果这个媒体名既不对应任何已有媒体名，也不对应任何已有媒体别名，则完成录入后标new_media_marker
					media_dict_b = self.add_media(media_dict) 
					if media_dict_b: #添加正确
						self.mark_media_done(media_dict)
				else: #如果这个媒体名对应任何已有媒体名或媒体别名，则直接标media_name_exists_marker或media_alias_exists_marker
					self.mark_media_done(media_dict)

if __name__ == '__main__':
	add_media = AddMedia()
	add_media.run()