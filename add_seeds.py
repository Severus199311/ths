#coding=utf-8
#浙报web端种子录入，不是电子报

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC 
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
import xlwt, xlrd
from xlutils.copy import copy
import time, re
from logger import logger


class AddSeeds():

	def __init__(self):
		
		self.workbook_title = 'seeds.xls' #待录入种子所在的文件
		self.sheet_name = 'Sheet1' #待录入种子所在的表名
		self.sheet_order = 0 #待录入种子所在的表序号
		#以上三项为最关键配置，忘记修改的话，excel表中的种子录入情况将出错，且无法修复

		self.if_successful_column_order = 1 #if_successful在第2列，所以是1
		self.if_already_exist_column_order = 11 #if_already_exist在第12，所以是11
		self.successful_marker = 1 #成功，则标记为1
		#以上三项也是重要配置，忘记修改的话，excel表中的种子录入情况将出错，但可以修复

		self.start_url = 'http://flashcms.10jqka.com.cn/default/index/index'
		self.start_url_2 = 'http://flashcms.10jqka.com.cn/seedmanager/html/index.html#/seed_edit'
		self.wait_time = 30 #selenium等待时长
		self.sleep_time = 0.3

		self.workbook_read = xlrd.open_workbook(self.workbook_title) #打开excel文件，用于读取数据
		self.sheet_data = self.workbook_read.sheet_by_name(self.sheet_name) #读取表数据
		self.keys_list = self.sheet_data.row_values(0) #获取列名，格式为list
		self.workbook_write = copy(self.workbook_read) #复制excel文件，用于写入
		self.processing_row = 0 #初始行序号为0

		self.browser = webdriver.Chrome() 
		self.wait = WebDriverWait(self.browser, self.wait_time)

	#打开种子录入页面
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
	def fetch_seed_undone(self):
		self.processing_row += 1
		value_list = self.sheet_data.row_values(self.processing_row)
		seed_dict = dict(zip(self.keys_list, value_list))
		if not seed_dict.get('if_successful'):
			#return seed_dict
			if seed_dict.get('css_path'): #必须满足两个条件：种子尚未录入，css_path已经找到
				return seed_dict
		return None

	#将新处理好的种子标记为1
	def mark_seed_done(self, seed_dict):
		sheet = self.workbook_write.get_sheet(self.sheet_order)
		sheet.write(self.processing_row, self.if_successful_column_order, self.successful_marker) #if_successful在第2列，所以是1
		self.workbook_write.save(self.workbook_title)
		logger.info(str(self.processing_row)+ ' ' + seed_dict.get('website') + ' ' + seed_dict.get('page') + ' 录入完成')

	#如果种子已经存在，将id标记
	def mark_seed_already_exists(self, seed_dict):
		sheet = self.workbook_write.get_sheet(self.sheet_order)
		sheet.write(self.processing_row, self.if_successful_column_order, self.successful_marker)
		seed_id = seed_dict.get('seed_id')
		sheet.write(self.processing_row, self.if_already_exist_column_order, seed_id)
		self.workbook_write.save(self.workbook_title)
		logger.info(str(self.processing_row)+ ' ' + seed_dict.get('website') + ' ' + seed_dict.get('page') + ' 已经存在' + ' 种子id：' + seed_id)

	def wait_and_get_element(self, scenario, css_path):
		if scenario == 'located':
			return self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, css_path)))
		if scenario == 'clickable':
			return self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, css_path)))

	#录入种子
	def add_seed(self, seed_dict):

		#录入种子的网站名
		website = self.wait_and_get_element('located', '#pane-web > div > form > div:nth-child(1) > div > div > div.el-col.el-col-20 > div > input')
		website.clear()
		website_text = ''
		for each in seed_dict.get('website'):
			if not each in [' ', '	', '\n']:
				website_text += each
		website.send_keys(website_text)

		#录入种子的页面名
		page = self.wait_and_get_element('located', '#pane-web > div > form > div:nth-child(2) > div > div > div.el-col.el-col-20 > div > input')
		page.clear()

		#录电子报要用到,其他都用不到
		if type(seed_dict.get('page')) == float:
			seed_dict['page'] = str(int(seed_dict.get('page')))

		page_text = ''
		for each in seed_dict.get('page'):
			if not each in [' ', '	', '\n']:
				page_text += each
		page.send_keys(page_text)

		#录入种子url
		url = self.wait_and_get_element('located', '#pane-web > div > form > div:nth-child(3) > div > div > div.el-col.el-col-20 > div > input')
		url.clear()
		url.send_keys(seed_dict.get('url'))

		#录入种子的文章css_path
		css_path = self.wait_and_get_element('located', '#pane-web > div > form > div:nth-child(4) > div > div > div.el-col.el-col-20 > div > input')
		css_path.clear()
		css_path.send_keys(seed_dict.get('css_path'))


		#录入种子分类
		category = self.wait_and_get_element('located', '#pane-web > div > form > div:nth-child(7) > div:nth-child(1) > div > div > div.el-select__tags > input')
		category.clear()
		category.send_keys(seed_dict.get('category'))
		time.sleep(self.sleep_time)
		target_category = self.wait_and_get_element('located', 'body > div.el-select-dropdown.el-popper.is-multiple > div.el-scrollbar')
		target_category.click()

		#点掉因上一步出现的框框，如果不点掉的话下面的元素可能定位不到
		author_id = self.wait_and_get_element('clickable', '#pane-web > div > form > div:nth-child(6) > div > div > div > div > input')
		author_id.click()

		
		#录入种子业务分组
		some_div = self.wait_and_get_element('located', '#pane-web > div > form > div:nth-child(7) > div:nth-child(2)')
		if some_div.find_element_by_css_selector('label').text == '种子业务分组':
			uses_list = some_div.find_elements_by_css_selector('div.el-form-item__content > div.el-checkbox-group > label')
			target_uses_list = seed_dict.get('use').split(',')
			for each in uses_list:
				if each.find_element_by_css_selector('span.el-checkbox__label').text in target_uses_list:
					target_use_button = each.find_element_by_css_selector('span.el-checkbox__input > span.el-checkbox__inner') 
					target_use_button.click() 
		else:
			logger.error('未定位到种子业务分组节点，程序退出。请检查css路径')
			sys.exit()

		#录入种子星级
		rating_level = str(int(seed_dict.get('rating')))
		rating_css_path = '#pane-web > div > form > div:nth-child(7) > div:nth-child(3) > div > div > span:nth-child({}) > i'.format(rating_level)
		rating = self.wait_and_get_element('clickable', rating_css_path)
		rating.click()

		##录入种子抓取频率，frequency_level 4,5,6,7分别表示1m,10m,30m,1h 
		frequency_level = str(int(seed_dict.get('frequency')))
		#frequency_css_path = 'body > div.el-select-dropdown.el-popper > div.el-scrollbar > div.el-select-dropdown__wrap.el-scrollbar__wrap > ul > li:nth-child({})'.format(frequency_level)
		frequency_css_path = 'body > div:nth-child(6) > div.el-scrollbar > div.el-select-dropdown__wrap.el-scrollbar__wrap > ul > li:nth-child({})'.format(frequency_level)
		#frequency_css_path = 'body > div.el-select-dropdown.el-popper > div.el-scrollbar > div.el-select-dropdown__wrap.el-scrollbar__wrap > ul > li:nth-child({})'.format(frequency_level)
		frequency = self.wait_and_get_element('clickable', '#pane-web > div > form > div:nth-child(8) > div > div > div.el-input.el-input--suffix > span')
		frequency.click()
		time.sleep(self.sleep_time)
		target_frequency = self.wait_and_get_element('located', frequency_css_path)
		target_frequency.click()

		add = self.wait_and_get_element('clickable', '#pane-web > div > div.footer > div > div:nth-child(3) > button:nth-child(1) > span')
		add.click()

		#捕捉提示框，如果提示种子已经存在，获取种子id
		message = self.wait_and_get_element('located', '.el-message__content').text
		if 'already exists' in message:
			seed_id = re.match('.*\[(.\d*)].*', message).group(1)
			seed_dict['seed_id'] = seed_id

			self.browser.refresh() #如果种子已经存在，需要刷新页面，以取消种子分类和业务分组

		time.sleep(self.sleep_time)
		return seed_dict

	def run(self):
		
		self.open_page()
		row_number = self.sheet_data.nrows - 1
		for row in range(row_number):
			seed_dict = self.fetch_seed_undone()
			if seed_dict:
				seed_dict = self.add_seed(seed_dict)
				if seed_dict.get('seed_id'):
					self.mark_seed_already_exists(seed_dict)
				else:
					self.mark_seed_done(seed_dict)

if __name__ == '__main__':
	add_seeds = AddSeeds()
	add_seeds.run()