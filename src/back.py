import bs4
import requests
import openpyxl as xl
import datetime
import pprint
import time

months = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь',
		  'Декабрь']


class Consultant:
	def __init__(self):
		self.quarters = {'Январь': [0, 0], 'Февраль': [0, 1], 'Март': [0, 2], 'Апрель': [1, 0], 'Май': [1, 1],
						 'Июнь': [1, 2], 'Июль': [2, 0], 'Август': [2, 1], 'Сентябрь': [2, 2], 'Ноябрь': [3, 0],
						 'Октябрь': [3, 1], 'Декабрь': [3, 2]}
		self.duration = {
			'40': 0,
			'36': 1,
			'24': 2
		}
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
		quarter = self.quarters[month][0]
		quarter_list = self.soup.find('div', class_='block-print').find_next_sibling().find_next_sibling(). \
			find_all('div', class_='row')  # Нашли divs
		quarter_list_result = [quarter_list[4]]  # Нашли первый квартал
		quarter_list = quarter_list[5:]  # Обрезали ввиду верстки
		for i in range(5, len(quarter_list), 6):
			quarter_list_result.append(quarter_list[i])
		standard_hours = \
			float(
				quarter_list_result[quarter].find_all_next('div', class_='col-md-3 col-xs-2')
				[self.quarters[month][1]].text \
					.split()[self.duration['40']].strip().replace(',', '.'))  # Магия и ты гений
		return standard_hours


class Excel:
	def __init__(self, path):
		self.workbook = xl.load_workbook(path)
		self.result_workbook = xl.Workbook()
		self.data = {}

	def get_jobs_and_div(self):
		sheet = self.workbook.active
		for row in sheet.iter_rows(2, sheet.max_row):
			self.data[row[1].value] = [row[2].value, row[3].value, row[5].value]
		return self.data

	def create_result_table(self, month, data):
		for i in range(data):  # Номера проекта
			sheet = self.result_workbook.create_sheet([data[i][0]])
			for i, j in zip(range(6),
							['Проект', 'ФИО', 'Должность', 'Отдел', 'Часы', 'Деньги']):
				sheet[0][i] = j

			sheet.merge_cells(f'A2', f'A{sheet.max_row - 1}')
			sheet['A2'] = data[i][0]

			sheet.merge_cells(f'A{sheet.max_row}', f'D{sheet.max_row}')
			sheet[f'A{sheet.max_row}'] = 'Итого'

		self.result_workbook.save(f'{month}_{datetime.datetime.today().year}.xlsx')

	def set_result_table(self, month, data):
		for i in range(data):  # Номер проекта
			sheet = self.result_workbook[data[i][0]]
			for j in range(1, 6):
				sheet[...][...] = data[1][...]  # Заполнение данных
			# sheet[f'E{sheet.max_row}'] =   # Итог часов на проект
			# sheet[f'F{sheet.max_row}'] =   # Итог выплаченных денег

		self.result_workbook.save(f'{month}_{datetime.datetime.today().year}.xlsx')


class Bitrix:
	def __init__(self, path):
		self.path = path
		self.workbook = xl.load_workbook(self.path)
		self.sheet = self.workbook.active

	def get_hours_worked(self):
		data = {}
		print(self.sheet.max_column - 3)
		for proj in range(3, self.sheet.max_column):
			for row in range(1, self.sheet.max_row):
				data[str(int(self.sheet[2][proj].value[:4]))] = dict([(self.sheet[i][1].value, self.sheet[i][proj].value) for i in range(3, self.sheet.max_row-1)])
		return data

	def get_projects(self):
		return [int(self.sheet[2][i].value[:4]) for i in range(3, self.sheet.max_row - 3)]

	def get_month_from_name(self):
		return self.path.split('\\')[-1].split('_')[0]


class Adesk:  # После получения доступа к данным, попробовать вытащить контрагента
	def __init__(self, api):
		# '415b6479c8df4d619ff3e957e6a262242f5c3fa6024744c2ac1ff532e8d76a1e'
		self.API = api
		self.data = {}

		response = requests.get(f'https://api.adesk.ru/v1/transactions?api_token={self.API}')
		while not response.ok:
			response = requests.get(f'https://api.adesk.ru/v1/projects?api_token={self.API}')
			time.sleep(1)
		self.income_json = response.json()

	def get_projects(self):
		response = requests.get(f'https://api.adesk.ru/v1/projects?api_token={self.API}')
		while not response.ok:
			response = requests.get(f'https://api.adesk.ru/v1/projects?api_token={self.API}')
			time.sleep(1)
		data_json = response.json()
		for i in range(len(data_json['projects'])):
			pprint.pprint(data_json)
			numbers_of_project = data_json['projects'][i]['name'][:4]
			id_of_project = data_json['projects'][i]['id']
			incomes = float(data_json['projects'][i]['income'])
			plan_incomes = float(data_json['projects'][i]['planIncome'])
			self.data[numbers_of_project] = (incomes, plan_incomes)
		return self.data

	def get_income_of_project_in_month(self, month, year, id_proj):
		income = 0
		start = datetime.datetime.strptime(f"1.{months.index(month) + 1}.{year}", "%d.%m.%Y")
		end = datetime.datetime.strptime(f"31.{months.index(month) + 1}.{year}", "%d.%m.%Y")
		for i in range(self.income_json['recordsTotal']):
			now = datetime.datetime.strptime(self.income_json['transactions'][i]['date'], "%d.%m.%Y")
			if (self.income_json['transactions'][i]['project']['id'] == id_proj) and (start <= now <= end):
				income += float(self.income_json['transactions'][i]['amount'])
		return income


class Accountant:
	def calculating_wages(self, worked_hours, standard_hours, salary_rate):
		return float(f"{(worked_hours / standard_hours) * salary_rate:.{2}f}")
