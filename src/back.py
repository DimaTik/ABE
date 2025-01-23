import bs4
import requests
import openpyxl
import datetime
import pprint
import time


class Consultant:
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

	def get_standard_hours(self, month):
		months = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
		quarters = {'Январь': 0, 'Февраль': 0, 'Март': 0, 'Апрель': 1, 'Май': 1, 'Июнь': 1, 'Июль': 2, 'Август': 2, 'Сентябрь': 2, 'Ноябрь': 3,	'Октябрь': 3, 'Декабрь': 3}
		quarter = quarters[month]
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
		standard_hours = quarter_list_result[quarter].find_all_next('div', class_='col-md-3 col-xs-2')[months.index(month)].text\
			.split()[duration['40']].strip() 	# Магия и ты гений
		return standard_hours


class Excel:
	def __init__(self, path):
		self.workbook = openpyxl.load_workbook(path)
		self.worksheet = self.workbook.active
		self.data = {}

	def get_jobs_and_div(self):
		for row in self.worksheet.iter_rows(2, self.worksheet.max_row):
			self.data[row[1].value] = [row[2].value, row[3].value, row[5].value]
		return self.data


class Bitrix:
	def __init__(self, path):
		self.workbook = openpyxl.load_workbook(path)
		self.worksheet = self.workbook.active
		self.data = {}

	def get_hours(self):
		for row in range(3, self.worksheet.max_row-1):
			self.data[self.worksheet[row][1].value] = [self.worksheet[row][2].value, [dict([(int(self.worksheet[2][i].value[:4]), self.worksheet[row][i].value)]) for i in range(3, self.worksheet.max_row-3)]]
		return self.data

	def get_projects_in_month(self):
		return [int(self.worksheet[2][i].value[:4]) for i in range(3, self.worksheet.max_row - 3)]


class Adesk: 	# После получения доступа к данным, попробовать вытащить контрагента
	def __init__(self, api):
		# '415b6479c8df4d619ff3e957e6a262242f5c3fa6024744c2ac1ff532e8d76a1e'
		self.API = api
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
