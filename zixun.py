#根据来源名和时间获取“资讯搜索”中的文章

import requests, xlwt, time
from xlutils.copy import copy
from logger import logger
from lxml import etree
from bs4 import BeautifulSoup

class ZIXUNNEWS():

	def __init__(self):

		self.start_date = time.strftime('%Y-%m-%d')
		self.end_date = time.strftime('%Y-%m-%d')
		self.page = 0
		self.next_page = True
		self.source = '%D2%F8%CA%C1%B2%C6%BE%AD' #'中文“银柿财经”不行'
		#资讯
		self.base_url = 'http://flashcms.10jqka.com.cn/input/formalnews/index/?title=&seq=&enter=&updater=&source={}&author=&hyname=&hyid=&conceptid=&newconceptid=&stockcode=&sclassname=&classid=&classname=&relate=0&remark=1&weight=&starttime={}&endtime={}&page={}'
		#同顺号
		#self.base_url = 'http://flashcms.10jqka.com.cn/entry/ContentServerAudit/index/?title=&author=%E6%B3%A1%E8%B4%A2%E7%BB%8F&status=&num=10&pid=&type=1&starttime={}%2000%3A00&refresh=0&endtime={}%2023%3A59&contentType=0&page={}'
		self.headers = {
			'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
			'Accept-Encoding':'gzip, deflate',
			'Accept-Language':'zh-CN,zh;q=0.9',
			'Connection':'keep-alive',
			'Cookie':'_gscu_86787713=66253878pio93z11; PHPSESSID=tft1f9dlha2uclqmm4oi42q9d4; _gscu_1064072334=66254158k0z8yi77; _gscbrs_1064072334=1; _gscbrs_86787713=1; Hm_lvt_e7cc17a9d20d160c9385a4c39e2ac5f2=1666319828; __bid_n=183f867a442541a0e14207; Hm_lpvt_e7cc17a9d20d160c9385a4c39e2ac5f2=1666320043; Hm_lvt_027bddd22c4ba55ba649646907df1ea9=1666337273; Hm_lpvt_027bddd22c4ba55ba649646907df1ea9=1666337273; Hm_lvt_c8006686bbff79fd9a8b78f87ba62d5b=1666339165; Hm_lpvt_c8006686bbff79fd9a8b78f87ba62d5b=1666339534; Hm_lvt_1251e2164368acb59cdf7d709b0485f6=1666578947; Hm_lpvt_1251e2164368acb59cdf7d709b0485f6=1666579281; Hm_lvt_fe76d250fc45bcdfc9267a4f6348f8d8=1666580863; Hm_lpvt_fe76d250fc45bcdfc9267a4f6348f8d8=1666580863; Hm_lvt_de22db29221daaaaeac0ee7b2ad1cabf=1666581706; Hm_lpvt_de22db29221daaaaeac0ee7b2ad1cabf=1666581706; Hm_lvt_a42709843eab10847462653598b51c65=1666589526; Hm_lpvt_a42709843eab10847462653598b51c65=1666589526; Hm_lvt_cc1c22da8a226f87e35e5763fff2177d=1666589949; Hm_lpvt_cc1c22da8a226f87e35e5763fff2177d=1666590103; Hm_lvt_1ff3578d5ba068a03106eda6549fd562=1666590830; Hm_lpvt_1ff3578d5ba068a03106eda6549fd562=1666591082; Hm_lvt_b30beac462776446fe5c2ba16bf74068=1666591728; Hm_lpvt_b30beac462776446fe5c2ba16bf74068=1666591728; Hm_lvt_78c58f01938e4d85eaf619eae71b4ed1=1666594445; Hm_lvt_2a607c847fd7c8e8c03f9a8e3ebf0489=1666595606; Hm_lpvt_2a607c847fd7c8e8c03f9a8e3ebf0489=1666595733; Hm_lvt_e8b9f0045112ac707e8fddf868995afb=1666608122; Hm_lpvt_e8b9f0045112ac707e8fddf868995afb=1666608130; Hm_lvt_d1c3cd5dff7cae342489efed8dea437f=1666608246; Hm_lpvt_d1c3cd5dff7cae342489efed8dea437f=1666608316; Hm_lvt_7104f48adc2385f9d7bf0aefcdd1d926=1666608787; Hm_lpvt_7104f48adc2385f9d7bf0aefcdd1d926=1666608851; __bid_n=1840d28bf7172423e04207; cid=4f3e56a654462f8ee2cd48482bac805f1666667865; Hm_lpvt_78c58f01938e4d85eaf619eae71b4ed1=1666674977; v=A41A5Qu-0kH963YXR53cejw4nKICasMaS5glEM8QzP2TC6Pcl7rRDNvuNXtc',
			'Host':'flashcms.10jqka.com.cn',
			'Upgrade-Insecure-Requests':'1',
			'User-Agent':'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36'
		}
		self.article_dicts_list =[]

	def get_zixun_news(self, url):

		response = requests.get(url, headers=self.headers)
		response.encoding = 'gbk'
		return response.text


	def parse_zixun_news(self, response_text):

		soup = BeautifulSoup(response_text, 'lxml')
		article_trs = soup.select('#newsTable tr')
		del article_trs[0]
		for article_tr in article_trs:
			article_dict = {}
			article_dict['url'] = article_tr.select('.oldUrl')[0].attrs.get('value')
			article_dict['title'] = article_tr.select('.c1')[0].get_text().replace('\n', '').strip(' ').replace('复制','')
			article_dict['time'] = article_tr.select('.c9')[0].get_text().replace(' ', '')
			self.article_dicts_list.append(article_dict)
		page_text = soup.select('.pages')[0].get_text()
		if not '下一页' in page_text:
			self.next_page = False


	def run(self):

		while self.next_page:
			self.page += 1
			url = self.base_url.format(self.source, self.start_date, self.end_date, str(self.page))
			response_text = self.get_zixun_news(url)
			self.parse_zixun_news(response_text)
		
		self.write_into_excel()

if __name__ == '__main__':

	zixun_news = ZIXUNNEWS()
	zixun_news.run()


	#编码问题。垃圾etree，以后再也不用了，以后都用BeautifulSoup
	"""
	def parse_zixun_news(self, response_text):

		html = etree.HTML(response_text)
		article_trs = html.xpath('.//tr[contains(@class,"contextMenu")]')

		for article_tr in article_trs:
			url = article_tr.xpath('./input[@class="oldUrl"]')[0]
			title = article_tr.xpath('./input[@class="title"]')[0]
			print(url)
			print(title)
			#print(etree.tostring(url).decode('utf-8'))
			#print(etree.tostring(title).decode('utf-8'))
	"""