#coding=utf-8
#对excel表格里的每个url进行请求，在excel表格中标记返回码,非异步
#当self.check_page_titles = True， 也用来获取页面标签
#当self.check_outdated = True， 也用来检测网站是否过期

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC 
from selenium.webdriver.support.wait import WebDriverWait
import xlwt, xlrd
from xlutils.copy import copy
import time, requests, re
from logger import logger


class CheckUrls():

	def __init__(self):
		self.workbook_title = 'urls_to_be_checked.xls' #待录入种子所在的文件
		self.sheet_name = 'Sheet1' #待录入种子所在的表
		self.workbook_read = xlrd.open_workbook(self.workbook_title) #打开excel文件，用于读取数据
		self.sheet_data = self.workbook_read.sheet_by_name(self.sheet_name) #读取表数据
		self.keys_list = self.sheet_data.row_values(0) #获取列名，格式为list
		self.workbook_write = copy(self.workbook_read) #复制excel文件，用于写入
		self.processing_row = 0
		self.status_code_column = 7 #status_code字段所在列
		#self.website_column = 2
		#self.page_column = 3
		#self.this_year_column = 13
		#self.previous_years_column = 14
		self.check_page_titles = False
		self.check_outdated = False
		self.check_status_codes = True

	#从excel表格中获取待检测url
	def fetch_url_unchecked(self):
		self.processing_row += 1
		value_list = self.sheet_data.row_values(self.processing_row)
		seed_dict = dict(zip(self.keys_list, value_list))
		if not seed_dict.get('status_code'):
			if not seed_dict.get('this_year'):
				return seed_dict
		else:
			return None

	#url检测完成后，录入status_code
	def mark_url_checked(self, seed_dict):
		sheet = self.workbook_write.get_sheet(0)

		if self.check_status_codes:
			sheet.write(self.processing_row, self.status_code_column, seed_dict.get('status_code'))
		if self.check_page_titles:
			sheet.write(self.processing_row, self.website_column, seed_dict.get('website'))
			sheet.write(self.processing_row, self.page_column, seed_dict.get('page'))
		if self.check_outdated:
			if seed_dict.get('if_outdated'):
				sheet.write(self.processing_row, self.this_year_column, seed_dict.get('if_outdated').get('this_year'))
				sheet.write(self.processing_row, self.previous_years_column, seed_dict.get('if_outdated').get('previous_years'))

		self.workbook_write.save(self.workbook_title)

	def get_page_title(self, text):
		try:
			text = text.replace('\n', '')
			if re.match('[\s\S]*?<title>(.*?)</title>', text):
				title = re.match('[\s\S]*?<title>(.*?)</title>', text).group(1)
			else:
				title = re.match('[\s\S]*?<TITLE>(.*?)</TITLE>', text).group(1)
		except AttributeError:
			return None

		for chara in ['-', '——', '—', '|', '>']:
			title = title.replace(chara, '_')
		title_list = title.split('_')
		title_list.reverse()

		if len(title_list) == 1:
			website = title_list[0]
			page = title_list[0]

		else:
			website = title_list.pop(0)
			page = '_'.join(title_list)

		website = website.strip('_')
		page = page.strip('_')
		return {'website': website, 'page': page}

	def get_year_info(self, text):
		year_number = 0
		for year in ['2004', '2005', '2006', '2007', '2008', '2009', '2010', '2011', '2012', '2013', '2013', '2014', '2015', '2016', '2017', '2018', '2019', '2020']:
			regex = re.compile(year)
			year_list = regex.findall(text)
			year_number += len(year_list)
		regex = re.compile('2021')
		year_list = regex.findall(text)
		return {'this_year': len(year_list), 'previous_years': str(year_number)}
		#return str(len(year_list)) + '/' + str(year_number)

	#检测url
	def check_url(self, seed_dict):
		url = seed_dict.get('url')
		if url:
			try:
				response = requests.get(url, timeout=3)
				status_code = response.status_code
			except Exception as e:
				response = None
				status_code = str(e.args)

			seed_dict['status_code'] = str(status_code)

			if status_code == 200:
				
				if self.check_page_titles:
					if not seed_dict.get('page'):

						try:
							charset = re.match('[\s\S]*?charset=(.*?)>', response.text).group(1)
						except AttributeError:
							charset = None

						if charset:
							if 'gb' in charset or 'GB' in charset:
								response.encoding = 'gbk'
							else:
								response.encoding = 'utf-8'
						else:
							response.encoding = 'utf-8'

						title = self.get_page_title(response.text)
						if title:
							seed_dict['website'] = title.get('website')
							seed_dict['page'] = title.get('page')

				if self.check_outdated:
					seed_dict['if_outdated'] = self.get_year_info(response.text)

		return seed_dict
	
	def run(self):
		row_number = self.sheet_data.nrows - 1
		for row in range(row_number):
			seed_dict = self.fetch_url_unchecked()
			if seed_dict:
				seed_dict = self.check_url(seed_dict)
				self.mark_url_checked(seed_dict)
				logger.info(str(int(seed_dict.get('order'))) +  ' ' + seed_dict.get('status_code'))

if __name__ == '__main__':
	check_urls = CheckUrls()
	check_urls.run()