#对已经存在的种子，更改业务分组、增加拓展抓取
#对应的excel表格叫seeds_to_be_edited.xls

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC 
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementNotInteractableException
import xlwt, xlrd
from xlutils.copy import copy
import time, re
from logger import logger

class EditSeeds():

	def __init__(self):

		self.edit_extension = False
		self.edit_uses = False
		self.edit_tags = True
		#以上三项必须事先检查
		
		self.workbook_title = 'seeds_to_be_edited.xls' #待录入种子所在的文件
		self.sheet_name = 'Sheet1' #待录入种子所在的表名
		self.sheet_order = 0 #待录入种子所在的表序号
		#以上三项为最关键配置，忘记修改的话，excel表中的种子录入情况将出错，且无法修复

		self.if_successful_column_order = 1 #if_successful在第2列，所以是1
		self.successful_marker = 1 #成功，则标记为1
		#以上三项也是重要配置，忘记修改的话，excel表中的种子录入情况将出错，但可以修复

		self.start_url = 'http://flashcms.10jqka.com.cn/default/index/index'
		self.start_url_2 = 'http://flashcms.10jqka.com.cn/seedmanager/html/index.html#/seed_central'
		self.wait_time = 30 #selenium等待时长
		self.sleep_time = 0.5

		self.workbook_read = xlrd.open_workbook(self.workbook_title) #打开excel文件，用于读取数据
		self.sheet_data = self.workbook_read.sheet_by_name(self.sheet_name) #读取表数据
		self.keys_list = self.sheet_data.row_values(0) #获取列名，格式为list
		self.workbook_write = copy(self.workbook_read) #复制excel文件，用于写入
		
		self.processing_row = 0 #初始行序号为0

		self.browser = webdriver.Chrome() 
		self.wait = WebDriverWait(self.browser, self.wait_time)

		self.refresh_when_20 = 0

	#打开种子录入页面
	def open_page(self):

		self.browser.get(self.start_url)
		self.wait.until(EC.alert_is_present())
		self.browser.switch_to.alert.accept()
		self.wait.until(EC.alert_is_present())
		self.browser.switch_to.alert.accept()
		#self.wait.until(EC.presence_of_element_located((By.ID, '_M7'))) #现在找不到_M7节点了，但貌似不要也没关系
		self.browser.get(self.start_url_2)
		self.browser.maximize_window()

	#从excel表格中获取待编辑种子信息
	def fetch_seed_undone(self):

		self.processing_row += 1
		value_list = self.sheet_data.row_values(self.processing_row)
		seed_dict = dict(zip(self.keys_list, value_list))
		if not seed_dict.get('if_successful'):
			return seed_dict
		return None

	#将新编辑好的种子标记为1
	def mark_seed_done(self, seed_dict):

		sheet = self.workbook_write.get_sheet(self.sheet_order)
		sheet.write(self.processing_row, self.if_successful_column_order, self.successful_marker) #if_successful在第2列，所以是1
		self.workbook_write.save(self.workbook_title)

	def wait_and_get_element(self, scenario, css_path):

		if scenario == 'located':
			return self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, css_path)))
		if scenario == 'clickable':
			return self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, css_path)))

	#编辑20次之后刷新，否则会崩溃
	def refresh(self):

		time.sleep(self.sleep_time * 4)
		self.browser.refresh()
		time.sleep(self.sleep_time * 4)

		self.refresh_when_20 = 0


	#修改业务分组
	def edit_seeds_2(self, seed_dict):

		uses_list = self.browser.find_elements_by_css_selector('#pane-web > div > form > div:nth-child(7) > div.el-form-item.is-required > div > div > label')

		for each in uses_list:
			if each.find_element_by_css_selector('span.el-checkbox__label').text in seed_dict.get('use').split(','): #如果该标签是excel中的目标标签
				if 'is-checked' in each.find_element_by_css_selector('span.el-checkbox__input').get_attribute('class'):
					confirm_editing = self.wait_and_get_element('clickable', '#pane-web > div > div.footer > div > div:nth-child(3) > button:nth-child(3) > span')
					confirm_editing.click()
					logger.info(str(self.processing_row)+ ' ' + seed_dict.get('seed_id') + ' ' + seed_dict.get('website') + '-'+ seed_dict.get('page') + ' ' + seed_id + ' 已有该业务分组')
					#logger.info(str(self.processing_row)+ ' ' + seed_dict.get('seed_id') + ' ' + seed_dict.get('key_word') + ' ' + seed_id + ' 已有该业务分组')
				else:
					target_use_button = each.find_element_by_css_selector('span.el-checkbox__input > span.el-checkbox__inner') 
					target_use_button.click()
					confirm_editing = self.wait_and_get_element('clickable', '#pane-web > div > div.footer > div > div:nth-child(3) > button:nth-child(3) > span')
					confirm_editing.click()
					logger.info(str(self.processing_row)+ ' ' + seed_dict.get('seed_id') + ' ' + seed_dict.get('website') + '-'+ seed_dict.get('page') + ' ' + seed_id + ' 将添加分组')
					#logger.info(str(self.processing_row)+ ' ' + seed_dict.get('seed_id') + ' ' + seed_dict.get('key_word') + ' ' + seed_id + ' 将添加分组')

	#修改种子分类
	def edit_seeds_3(self, seed_dict):

		old_tag_button = self.wait_and_get_element('clickable', '#pane-web > div > form > div:nth-child(7) > div:nth-child(1) > div > div > div.el-select__tags > span > span > i')
		old_tag_button.click()
		time.sleep(self.sleep_time)

		tag_box = self.wait_and_get_element('located', '#pane-web > div > form > div:nth-child(7) > div:nth-child(1) > div > div > div.el-select__tags > input')
		tag_box.send_keys(seed_dict.get('category'))

		target_category = self.wait_and_get_element('located', 'body > div.el-select-dropdown.el-popper.is-multiple > div.el-scrollbar')
		target_category.click()

		#点掉因上一步出现的框框
		author_id = self.wait_and_get_element('clickable', '#pane-web > div > form > div:nth-child(6) > div > div > div > div > input')
		author_id.click()


	#根据种子id或关键词搜索种子
	def search_seeds(self, seed_dict):

		#不刷新会卡死
		if self.refresh_when_20 == 20:
			self.refresh()


		"""
		#根据source_id搜索。seed_id字段有可能是float类型。如果是，就换成string
		seed_id = seed_dict.get('seed_id')
		if isinstance(seed_id, float):
			seed_id = str(int(seed_id))
			seed_dict['seed_id'] = seed_id

		sourceid_button = self.wait_and_get_element('located', '#app > div.right-content > div > div.search > div:nth-child(2) > input')
		sourceid_button.clear()
		sourceid_button.send_keys(seed_dict.get('seed_id'))
		"""

		#根据关键词搜索。必须保证搜出来结果唯一
		#key_word = seed_dict.get('key_word')
		key_word = seed_dict.get('website') + '-' + seed_dict.get('page')

		if "(" in key_word:
			logger.info(str(self.processing_row)+ ' ' + seed_dict.get('seed_id') + ' ' + seed_dict.get('website') + '-'+ seed_dict.get('page') + ' 有（），搜不出来')
			#logger.info(str(self.processing_row)+ ' ' + seed_dict.get('seed_id') + ' ' + seed_dict.get('key_word') + ' ' + seed_id + ' 有（），搜不出来')
			return False
		
		key_word_button = self.wait_and_get_element('located', '#app > div.right-content > div > div.search > div:nth-child(1) > input')
		key_word_button.clear()
		key_word_button.send_keys(key_word)

		search_button = self.wait_and_get_element('clickable', '#app > div.right-content > div > div.search > button')
		search_button.click()
		time.sleep(self.sleep_time)
		
		#self.wait_and_get_element('located', '#app > div.right-content > div > div.seed-table > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--enable-row-transition.el-table--mini > div.el-table__body-wrapper.is-scrolling-none > table > tbody > tr:nth-child(1)')
		self.wait_and_get_element('located', '#app > div.right-content > div > div.seed-table > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--mini > div.el-table__header-wrapper > table > thead')
		page_seeds_list = self.browser.find_elements_by_css_selector('#app > div.right-content > div > div.seed-table > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--enable-row-transition.el-table--mini > div.el-table__body-wrapper.is-scrolling-none > table > tbody > tr')

		#有5此搜索机会，当页面只有一条种子且其名称与excel表中名称一致时，所搜素的种子已经出来
		times = 0
		while len(page_seeds_list) != 1:
		#while len(page_seeds_list) != 1 or page_seeds_list[0].find_elements_by_css_selector('td:nth-child(2) > div > a')[0].text != key_word:
			if times > 4:
				return False
			times += 1
			search_button.click()
			time.sleep(self.sleep_time)
			page_seeds_list = self.browser.find_elements_by_css_selector('#app > div.right-content > div > div.seed-table > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--enable-row-transition.el-table--mini > div.el-table__body-wrapper.is-scrolling-none > table > tbody > tr')
		
		#logger.info(str(len(page_seeds_list)) + ' ' + page_seeds_list[0].find_elements_by_css_selector('td:nth-child(2) > div > a')[0].text)
		
		edit_button = self.browser.find_element_by_css_selector('#app > div.right-content > div > div.seed-table > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--enable-row-transition.el-table--mini > div.el-table__fixed-right > div.el-table__fixed-body-wrapper > table > tbody > tr > td.is-center > div > div > div:nth-child(1) > button')
		edit_button.click()
		time.sleep(self.sleep_time)

		return True

	def confirm_editing(self, seed_dict):

		confirm_editing_button = self.wait_and_get_element('clickable', '#pane-web > div > div.footer > div > div.el-col.el-col-2 > button:nth-child(3) > span')                                                
		confirm_editing_button.click()

		logger.info(str(self.processing_row)+ ' ' + seed_dict.get('seed_id') + ' ' + seed_dict.get('website') + '-'+ seed_dict.get('page') + ' 编辑完成')
		#logger.info(str(self.processing_row)+ ' ' + seed_dict.get('seed_id') + ' ' + seed_dict.get('key_word') + ' 编辑完成')

		self.refresh_when_20 += 1

	def run(self):

		self.open_page()

		row_number = self.sheet_data.nrows - 1
		for row in range(row_number):
			seed_dict = self.fetch_seed_undone()
			if seed_dict:

				if self.search_seeds(seed_dict):

					if self.edit_tags:
						self.edit_seeds_3(seed_dict)

					if self.edit_uses:
						self.edit_seeds_2(seed_dict)

					self.confirm_editing(seed_dict)

					self.mark_seed_done(seed_dict)

if __name__ == '__main__':
	edit_seeds = EditSeeds()
	edit_seeds.run()


"""
	#根据seed_id，对电子报种子添加“拓展抓取”, 或修改业务分组
	def edit_seeds(self, seed_dict):

		if self.refresh_when_20 == 20:
			self.refresh()

		#seed_id字段有可能是float类型。如果是，就换成string
		seed_id = seed_dict.get('seed_id')
		if isinstance(seed_id, float):
			seed_id = str(int(seed_id))
			seed_dict['seed_id'] = seed_id

		sourceid_button = self.wait_and_get_element('located', '#app > div.right-content > div > div.search > div:nth-child(2) > input')
		sourceid_button.clear()
		sourceid_button.send_keys(seed_dict.get('seed_id'))

		search_button = self.wait_and_get_element('clickable', '#app > div.right-content > div > div.search > button')
		search_button.click()
		time.sleep(self.sleep_time)
		

		#点击后续反应一段时间，所搜索的种子才能出来。判断条件：当页面只有一条种子且其名称与excel表中名称一致时，所搜素的种子已经出来
		self.wait_and_get_element('located', '#app > div.right-content > div > div.seed-table > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--enable-row-transition.el-table--mini > div.el-table__body-wrapper.is-scrolling-none > table > tbody > tr:nth-child(1)')
		page_seeds_list = self.browser.find_elements_by_css_selector('#app > div.right-content > div > div.seed-table > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--enable-row-transition.el-table--mini > div.el-table__body-wrapper.is-scrolling-none > table > tbody > tr')
		
		while len(page_seeds_list) > 1 or page_seeds_list[0].find_elements_by_css_selector('td:nth-child(2) > div > a')[0].text != seed_dict.get('seed_title'):
			search_button.click()
			time.sleep(self.sleep_time)
			page_seeds_list = self.browser.find_elements_by_css_selector('#app > div.right-content > div > div.seed-table > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--enable-row-transition.el-table--mini > div.el-table__body-wrapper.is-scrolling-none > table > tbody > tr')
		
		logger.info(str(len(page_seeds_list)) + ' ' + page_seeds_list[0].find_elements_by_css_selector('td:nth-child(2) > div > a')[0].text)
		
		edit_button = self.browser.find_element_by_css_selector('#app > div.right-content > div > div.seed-table > div.el-table.el-table--fit.el-table--striped.el-table--border.el-table--enable-row-transition.el-table--mini > div.el-table__fixed-right > div.el-table__fixed-body-wrapper > table > tbody > tr > td.is-center > div > div > div:nth-child(1) > button')
		edit_button.click()
		time.sleep(self.sleep_time)

		if self.edit_extension:

			self.wait_and_get_element('located', '#pane-web > div > form > div:nth-child(7) > div.el-form-item.is-required > div > div > label.el-checkbox.is-checked') 
			extension_input = self.wait_and_get_element('located', '#pane-web > div > form > div:nth-child(9) > div > div > div > div > input')
			extension_input.clear()

			try:
				extension_input.clear()
			except ElementNotInteractableException: #不知道为什么，有时候会进入到总表的第一个种子。如果这个种子不是web，就会出现ElementNotInteractableException；如果是web，那就真的会录错
				self.refresh()
				return None

			extension_input.send_keys('{"needPageList": 1}')

		if self.edit_uses:

			uses_list = self.browser.find_elements_by_css_selector('#pane-web > div > form > div:nth-child(7) > div.el-form-item.is-required > div > div > label')

			for each in uses_list:
				use = each.find_element_by_css_selector('span.el-checkbox__label').text
				if use in self.target_uses_list:
					if 'is-checked' in each.find_element_by_css_selector('span.el-checkbox__input').get_attribute('class'):
						logger.info(str(self.processing_row)+ ' ' + seed_dict.get('seed_title') + ' ' + seed_id + ' 已有业务分组 ' + use)
					else:
						target_use_button = each.find_element_by_css_selector('span.el-checkbox__input > span.el-checkbox__inner') 
						target_use_button.click()
						logger.info(str(self.processing_row)+ ' ' + seed_dict.get('seed_title') + ' ' + seed_id + ' 将添加分组 ' + use)

		time.sleep(self.sleep_time)

		confirm_editing = self.wait_and_get_element('clickable', '#pane-web > div > div.footer > div > div.el-col.el-col-2 > button:nth-child(3) > span')                                                
		confirm_editing.click()
		logger.info(str(self.processing_row)+ ' ' + seed_id + ' 编辑完成')

		self.refresh_when_20 += 1

		return seed_dict
"""