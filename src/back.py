import bs4
import requests
import openpyxl
import datetime
import pprint
import time


class Parser:
	def __init__(self):
		headers = {
			'Accept': "*/*",
			'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 OPR/115.0.0.0 (Edition Yx 05)"
		}
		year = datetime.date.today().year
		url = f'https://www.consultant.ru/law/ref/calendar/proizvodstvennye/{year}'
		req = requests.get(url, headers)
		with open(fr'C:\Users\dima2\AppData\Local\Temp\calendar_{year}', 'w', encoding='utf-8') as file:
			file.write(req.text)
		with open(fr'C:\Users\dima2\AppData\Local\Temp\calendar_{year}', 'r', encoding='utf-8') as file:
			src = file.read()

		self.soup = bs4.BeautifulSoup(src, 'lxml')

	def get_standard_hours(self, quarter, month, week_duration):
		duration = {
			'40': 0,
			'36': 1,
			'24': 2
		}
		quarter_list = self.soup.find('div', class_='block-print').find_next_sibling().find_next_sibling().\
			find_all('div', class_='row') 			# Нашли divs
		quarter_list_result = [quarter_list[4]] 	# Нашли первый квартал
		quarter_list = quarter_list[5:]				# Обрезали ввиду верстки
		for i in range(5, len(quarter_list), 6):
			quarter_list_result.append(quarter_list[i])
		standard_hours = quarter_list_result[quarter].find_all_next('div', class_='col-md-3 col-xs-2')[month].text\
			.split()[duration[week_duration]].strip() 	# Магия и ты гений
		# print(standard_hours)
		return standard_hours


class Excel:
	def __init__(self, path):
		self.workbook = openpyxl.load_workbook(path)
		self.worksheet = self.workbook.active

	def get_jobs_and_div(self):
		dict = {}
		for i in range(0, self.worksheet.max_column):
			for col in self.worksheet.iter_rows(2, self.worksheet.max_row):
				dict[col[1].value] = [col[2].value, col[3].value]
		return dict


class Adesk: 	# После получения доступа к данным, попробовать вытащить контрагента
	def __init__(self):
		self.API = '415b6479c8df4d619ff3e957e6a262242f5c3fa6024744c2ac1ff532e8d76a1e'
		self.data = {}

	def get_project(self):
		response = requests.get(f'https://api.adesk.ru/v1/projects?api_token={self.API}')
		while not response.ok:
			response = requests.get(f'https://api.adesk.ru/v1/projects?api_token={self.API}')
			time.sleep(2)
		data_json = response.json()
		for project in range(len(data_json['projects'])):
			pprint.pprint(data_json)
			numbers_of_project = data_json['projects'][project]['name'][:4]
			incomes = float(data_json['projects'][project]['income'])
			plan_incomes = float(data_json['projects'][project]['planIncome'])
			self.data[numbers_of_project] = (incomes, plan_incomes)
		return self.data
