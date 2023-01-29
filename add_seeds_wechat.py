#coding=utf-8
#微信种子录入

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC 
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
import xlwt, xlrd
from xlutils.copy import copy
import time, re, sys,json
from logger import logger

class AddSeedsWechatTemp():

	def __init__(self):

		self.workbook_title = 'seeds_wechat_temp.xls' #待录入种子所在的文件
		self.sheet_name = 'Sheet1' #待录入种子所在的表名
		self.sheet_order = 0 #待录入种子所在的表序号
		#以上三项为最关键配置，忘记修改的话，excel表中的种子录入情况将出错，且无法修复

		self.if_successful_column_order = 1 #if_successful在第2列，所以是1
		self.if_already_exist_column_order = 7 #if_already_exist在第13，所以是12
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

		wechat_page = self.wait_and_get_element('clickable', '#tab-WECHAT')
		wechat_page.click()

	def add_seed_wechat(self, seed_dict):

		if not seed_dict.get('wechat_id'):
			logger.error('缺少微信ID')
			return None
		if not seed_dict.get('wechat_name'):
			logger.error('缺少微信名称')
			return None

		#以下这种方式去找节点最稳妥，只能报错，不可能录错——除非表格本身就有错

		some_div = self.wait_and_get_element('located', '#pane-WECHAT > div > form > div:nth-child(1)')
		if some_div.find_element_by_css_selector('label').text == '微信名称':
			account_name_input = some_div.find_element_by_css_selector('div > div.el-input > input')
			account_name_input.clear()
			account_name_input.send_keys(seed_dict.get('wechat_name'))
		else:
			logger.error('未定位到微信名称节点，程序退出。请检查css路径')

		some_div = self.wait_and_get_element('located', '#pane-WECHAT > div > form > div:nth-child(2)')
		if some_div.find_element_by_css_selector('label').text == '微信ID':
			account_id_input = some_div.find_element_by_css_selector('div > div.el-input > input')
			account_id_input.clear()
			account_id_input.send_keys(seed_dict.get('wechat_id'))
		else:
			logger.error('未定位到微信ID节点，程序退出。请检查css路径')

		
		categories_list = seed_dict.get('seed_category').split(',')
		some_div = self.wait_and_get_element('located', '#pane-WECHAT > div > form > div:nth-child(4) > div:nth-child(1)')

		if some_div.find_element_by_css_selector('label').text == '种子分类':
			seed_category_input = some_div.find_element_by_css_selector('div > div > div.el-select__tags > input')
			seed_category_input.clear()

			for each in categories_list:
				seed_category_input.send_keys(each)
				time.sleep(self.sleep_time)
				target_categories_list = self.browser.find_elements_by_css_selector('body > div.el-select-dropdown.el-popper.is-multiple > div.el-scrollbar > div.el-select-dropdown__wrap.el-scrollbar__wrap > ul > li')
				for target_category in target_categories_list:
					if target_category.text == each:
						target_category.click()
						time.sleep(self.sleep_time)

		else:
			logger.error('未定位到种子分类节点，程序退出。请检查css路径')
			sys.exit()

		#点掉因上一步出现的框框
		author_id = self.wait_and_get_element('clickable', '#pane-WECHAT > div > form > div:nth-child(3) > div > div > input')
		author_id.click() 

		some_div = self.wait_and_get_element('located', '#pane-WECHAT > div > form > div:nth-child(4) > div:nth-child(2)')
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

		some_div = self.wait_and_get_element('located', '#pane-WECHAT > div > form > div:nth-child(4) > div:nth-child(3)')
		if some_div.find_element_by_css_selector('label').text == '种子星级':

			rating_css_path = 'div > div > span:nth-child({}) > i'.format(str(int(seed_dict.get('rating'))))
			rating = some_div.find_element_by_css_selector(rating_css_path)
			rating.click()

		else:
			logger.error('未定位到种子星级节点，程序退出。请检查css路径')
			sys.exit()

		add = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#pane-WECHAT > div > div > button > span')))
		add.click()
		#time.sleep(self.sleep_time)

		message = self.wait_and_get_element('located', '.el-message__content').text
		if 'already exists' in message:
			seed_id = re.match('.*\[(.\d*)].*', message).group(1)
			seed_dict['seed_id'] = seed_id

			#页面刷新，另辟蹊径，就不用把下面的业务分组和种子分类点掉了
			self.browser.refresh()
			wechat_page = self.wait_and_get_element('clickable', '#tab-WECHAT')
			wechat_page.click()

		time.sleep(self.sleep_time)
		return seed_dict


	#这里的wechat_account是一个字典格式，不是wechat_name和wechat_id各自一个字段
	def add_seed_wechat_2(self, seed_dict):
		wechat_account_json = seed_dict.get('wechat_name')
		
		try:
			wechat_account_list = json.loads(wechat_account_json)
		except json.decoder.JSONDecodeError:
			return None

		if len(wechat_account_list) != 1:
			return  None

		wechat_account_dict = wechat_account_list[0]

		some_div = self.wait_and_get_element('located', '#pane-WECHAT > div > form > div:nth-child(1)')
		if some_div.find_element_by_css_selector('label').text == '微信名称':
			account_name_input = some_div.find_element_by_css_selector('div > div.el-input > input')
			account_name_input.clear()
			account_name_input.send_keys(wechat_account_dict.get('wechat_name'))
		else:
			logger.error('未定位到微信名称节点，程序退出。请检查css路径')

		some_div = self.wait_and_get_element('located', '#pane-WECHAT > div > form > div:nth-child(2)')
		if some_div.find_element_by_css_selector('label').text == '微信ID':
			account_id_input = some_div.find_element_by_css_selector('div > div.el-input > input')
			account_id_input.clear()
			account_id_input.send_keys(wechat_account_dict.get('wechat_id'))
		else:
			logger.error('未定位到微信ID节点，程序退出。请检查css路径')

		
		categories_list = ['非上市重要公司']
		some_div = self.wait_and_get_element('located', '#pane-WECHAT > div > form > div:nth-child(4) > div:nth-child(1)')

		if some_div.find_element_by_css_selector('label').text == '种子分类':
			seed_category_input = some_div.find_element_by_css_selector('div > div > div.el-select__tags > input')
			seed_category_input.clear()

			for each in categories_list:
				seed_category_input.send_keys(each)
				time.sleep(self.sleep_time)
				target_categories_list = self.browser.find_elements_by_css_selector('body > div.el-select-dropdown.el-popper.is-multiple > div.el-scrollbar > div.el-select-dropdown__wrap.el-scrollbar__wrap > ul > li')
				for target_category in target_categories_list:
					if target_category.text == each:
						target_category.click()
						time.sleep(self.sleep_time)

		else:
			logger.error('未定位到种子分类节点，程序退出。请检查css路径')
			sys.exit()

		#点掉因上一步出现的框框
		author_id = self.wait_and_get_element('clickable', '#pane-WECHAT > div > form > div:nth-child(3) > div > div > input')
		author_id.click() 

		some_div = self.wait_and_get_element('located', '#pane-WECHAT > div > form > div:nth-child(4) > div:nth-child(2)')
		if some_div.find_element_by_css_selector('label').text == '种子业务分组':
			uses_list = some_div.find_elements_by_css_selector('div.el-form-item__content > div.el-checkbox-group > label')
			target_uses_list = ['资讯', '快讯']
			for each in uses_list:
				if each.find_element_by_css_selector('span.el-checkbox__label').text in target_uses_list:
					target_use_button = each.find_element_by_css_selector('span.el-checkbox__input > span.el-checkbox__inner') 
					target_use_button.click() 
		else:
			logger.error('未定位到种子业务分组节点，程序退出。请检查css路径')
			sys.exit()

		some_div = self.wait_and_get_element('located', '#pane-WECHAT > div > form > div:nth-child(4) > div:nth-child(3)')
		if some_div.find_element_by_css_selector('label').text == '种子星级':

			rating_css_path = 'div > div > span:nth-child({}) > i'.format(str(3))
			rating = some_div.find_element_by_css_selector(rating_css_path)
			rating.click()

		else:
			logger.error('未定位到种子星级节点，程序退出。请检查css路径')
			sys.exit()

		add = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#pane-WECHAT > div > div > button > span')))
		add.click()
		#time.sleep(self.sleep_time)

		message = self.wait_and_get_element('located', '.el-message__content').text
		if 'already exists' in message:
			seed_id = re.match('.*\[(.\d*)].*', message).group(1)
			seed_dict['seed_id'] = seed_id

			#页面刷新，另辟蹊径，就不用把下面的业务分组和种子分类点掉了
			self.browser.refresh()
			wechat_page = self.wait_and_get_element('clickable', '#tab-WECHAT')
			wechat_page.click()

		time.sleep(self.sleep_time)
		return seed_dict

	def wait_and_get_element(self, scenario, css_path):
		if scenario == 'located':
			return self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, css_path)))
		if scenario == 'clickable':
			return self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, css_path)))

	#从excel表格中获取待录入种子信息
	def fetch_seed_undone(self):
		self.processing_row += 1
		value_list = self.sheet_data.row_values(self.processing_row)
		seed_dict = dict(zip(self.keys_list, value_list))
		if not seed_dict.get('if_successful'): #必须满足的条件：种子尚未录入
			return seed_dict
		return None

	#将新处理好的种子标记为1
	def mark_seed_done(self, seed_dict):
		sheet = self.workbook_write.get_sheet(self.sheet_order)
		sheet.write(self.processing_row, self.if_successful_column_order, self.successful_marker) #if_successful在第2列，所以是1
		self.workbook_write.save(self.workbook_title)
		logger.info(str(self.processing_row)+ ' ' + seed_dict.get('wechat_name') + ' 录入完成')

	#如果种子已经存在，将id标记
	def mark_seed_already_exists(self, seed_dict):
		sheet = self.workbook_write.get_sheet(self.sheet_order)
		sheet.write(self.processing_row, self.if_successful_column_order, self.successful_marker)
		seed_id = seed_dict.get('seed_id')
		sheet.write(self.processing_row, self.if_already_exist_column_order, seed_id)
		self.workbook_write.save(self.workbook_title)
		logger.info(str(self.processing_row)+ ' ' + seed_dict.get('wechat_name') + ' 已经存在' + ' 种子id：' + seed_id)

	def run(self):
		self.open_page()
		row_number = self.sheet_data.nrows - 1
		for row in range(row_number):
			seed_dict = self.fetch_seed_undone()
			if seed_dict:
				#seed_dict = self.add_seed_wechat_2(seed_dict)
				seed_dict = self.add_seed_wechat(seed_dict)
				if seed_dict:
					if seed_dict.get('seed_id'):
						self.mark_seed_already_exists(seed_dict)
					else:
						self.mark_seed_done(seed_dict)

if __name__ == '__main__':
	add_seeds_wechat_temp = AddSeedsWechatTemp()
	add_seeds_wechat_temp.run()