#根据媒体名和站点名搜索媒体，然后进行编辑

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC 
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
import xlwt, xlrd
from xlutils.copy import copy
import time, re, json, requests
from logger import logger

class EditMedia():

	def __init__(self):

		self.workbook_title = 'media_to_be_edited.xls' #待编辑媒体所在的文件
		self.sheet_name = 'Sheet1' #待编辑媒体所在的表名
		self.sheet_order = 0 #待编辑媒体所在的表序号
		self.tag = '机构分析'

		self.if_successful_column_order = 1 #if_successful在第2列，所以是1
		self.successful_marker = 1 #成功，则标记为1

		self.start_url = 'http://flashcms.10jqka.com.cn/default/index/index'
		self.wait_time = 10 #selenium等待时长

		self.workbook_read = xlrd.open_workbook(self.workbook_title) #打开excel文件，用于读取数据
		self.sheet_data = self.workbook_read.sheet_by_name(self.sheet_name) #读取表数据
		self.keys_list = self.sheet_data.row_values(0) #获取列名，格式为list
		self.workbook_write = copy(self.workbook_read) #复制excel文件，用于写入
		self.processing_row = 0 #初始行序号为0

		self.browser = webdriver.Chrome() 
		self.wait = WebDriverWait(self.browser, self.wait_time)

	def this_sleep(self):
		time.sleep(0.4)

	def wait_and_get_element(self, scenario, css_path):
		if scenario == 'located':
			return self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, css_path)))
		if scenario == 'clickable':
			return self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, css_path)))

	#打开媒体编辑页面
	def open_page(self):
		self.browser.get(self.start_url)
		self.wait.until(EC.alert_is_present())
		self.browser.switch_to.alert.accept()
		self.wait.until(EC.alert_is_present())
		self.browser.switch_to.alert.accept()

		button_a = self.wait_and_get_element('clickable', '#_M1 > a')
		button_a.click()

		button_b = self.wait_and_get_element('clickable', '#s_media_list > a')
		button_b.click()

		self.browser.maximize_window()

	#从excel表格中获取待编辑媒体信息
	def fetch_media_undone(self):
		self.processing_row += 1
		value_list = self.sheet_data.row_values(self.processing_row)
		media_dict = dict(zip(self.keys_list, value_list))
		if not media_dict.get('if_successful'):
			logger.info(str(self.processing_row)+ ' ' + media_dict.get('media_name') + ' 准备编辑')
			return media_dict
		return None

	def edit_media(self, media_dict):

		self.wait_and_get_element('located', '#rightMain')
		self.browser.switch_to.frame('rightMain')

		media_name_input = self.wait_and_get_element('located', '#name')
		media_name_input.clear()
		media_name_input.send_keys(media_dict.get('media_name'))

		self.this_sleep()

		search_button = self.wait_and_get_element('clickable', '#btn-search')
		search_button.click()

		self.this_sleep()

		#有的媒体已经被删了
		try:
			self.wait_and_get_element('located', '#tbl-list > tbody > tr')
		except TimeoutException:
			media_name_input.clear()
			self.browser.switch_to.default_content()
			return None

		media_list = self.browser.find_elements_by_css_selector('#tbl-list > tbody > tr')

		"""
		while len(media_list) > 40:
			search_button.click()
			self.this_sleep()
			media_list = self.browser.find_elements_by_css_selector('#tbl-list > tbody > tr')
		"""	

		#超过40个，说明搜索没成功
		if len(media_list) != 1:
			media_name_input.clear()
			self.browser.switch_to.default_content()
			return None

		for media in  media_list:
			if media.find_elements_by_css_selector('td.l-name')[0].text == media_dict.get('media_name') and media.find_elements_by_css_selector('td.l-tag')[0].text == self.tag:
				
				try:

					#编辑
					media.find_elements_by_css_selector('td:last-child > button.btn.btn-primary.btn-edit')[0].click()

					self.wait_and_get_element('located', '#screenhost')

					tag_input = self.wait_and_get_element('clickable', '#tag')
					tag_input.click()

					self.this_sleep()

					tag_blank = self.wait_and_get_element('clickable', '#tag > option:first-child')
					tag_blank.click()

					self.this_sleep()

					submit_button = self.wait_and_get_element('clickable', '#btn-fd')
					submit_button.click()

					self.wait.until(EC.alert_is_present())
					self.browser.switch_to.alert.accept()
					
					self.this_sleep()
					self.browser.switch_to.default_content()

					return media_dict

				except TimeoutException:

					self.browser.switch_to.default_content()
					button_b = self.wait_and_get_element('clickable', '#s_media_list > a')
					button_b.click()
					self.this_sleep()
					return None


		media_name_input.clear()
		self.browser.switch_to.default_content()
		
		return None

	def mark_media_done(self, media_dict):
		sheet = self.workbook_write.get_sheet(self.sheet_order)
		sheet.write(self.processing_row, self.if_successful_column_order, self.successful_marker) #该标记其实有1，2，3，分别对应录入成功、媒体名存在、媒体别名存在
		self.workbook_write.save(self.workbook_title)
		logger.info(str(self.processing_row)+ ' ' + media_dict.get('media_name') + ' 编辑成功')

	def run(self):

		self.open_page()
		row_number = self.sheet_data.nrows - 1

		for row in range(row_number):
			media_dict = self.fetch_media_undone()
			if media_dict: #已经标注过的，不会再进行下面代码
				media_dict = self.edit_media(media_dict)
				if media_dict:
					self.mark_media_done(media_dict)

if __name__ == '__main__':
	edit_media = EditMedia()
	edit_media.run()