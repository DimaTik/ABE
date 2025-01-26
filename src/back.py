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
		months = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь',
				  'Ноябрь', 'Декабрь']
		quarters = {'Январь': 0, 'Февраль': 0, 'Март': 0, 'Апрель': 1, 'Май': 1, 'Июнь': 1, 'Июль': 2, 'Август': 2,
					'Сентябрь': 2, 'Ноябрь': 3, 'Октябрь': 3, 'Декабрь': 3}
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
		standard_hours = \
			float(
				quarter_list_result[quarter].find_all_next('div', class_='col-md-3 col-xs-2')[months.index(month)].text \
				.split()[duration['40']].strip())  # Магия и ты гений
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
		for proj in data[...]: 	# Номер проекта
			sheet = self.result_workbook.create_sheet([str(proj)])
			for i, j in zip(range(6),
							['Проект', 'ФИО', 'Должность', 'Отдел', 'Часы', 'Деньги']):
				sheet[0][i] = j

			sheet.merge_cells(f'A2', f'A{sheet.max_row - 1}')
			sheet['A2'] = proj

			sheet.merge_cells(f'A{sheet.max_row}', f'D{sheet.max_row}')
			sheet[f'A{sheet.max_row}'] = 'Итого'

		self.result_workbook.save(f'{month}_{datetime.datetime.today().year}.xlsx')

	def set_result_table(self, month, data):
		for proj in data[...]: 	# Номер проекта
			sheet = self.result_workbook[str(proj)]
			for i in range(1, 6):
				sheet[...][...] = data[...] 	# Заполнение данных
			sheet[f'E{sheet.max_row}'] = data[...]  # Итог часов на проект
			sheet[f'F{sheet.max_row}'] = data[...]  # Итог выплаченных денег

		self.result_workbook.save(f'{month}_{datetime.datetime.today().year}.xlsx')


class Bitrix:
	def __init__(self, path):
		self.path = path
		self.workbook = xl.load_workbook(self.path)
		self.sheet = self.workbook.active

	def get_hours_worked(self):
		data = {}
		for row in range(3, self.sheet.max_row - 1):
			data[self.sheet[row][1].value] = [self.sheet[row][2].value, [
				dict([(int(self.sheet[2][i].value[:4]), self.sheet[row][i].value)]) for i in
				range(3, self.sheet.max_row - 3)]]
		return data

	def get_projects_in_month(self):
		return [int(self.sheet[2][i].value[:4]) for i in range(3, self.sheet.max_row - 3)]

	def get_month_from_name(self):
		return self.path.split('\\')[-1].split('_')[0]


class Adesk: 	# После получения доступа к данным, попробовать вытащить контрагента
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
		start = datetime.datetime.strptime(f"1.{months.index(month)+1}.{year}", "%d.%m.%Y")
		end = datetime.datetime.strptime(f"31.{months.index(month)+1}.{year}", "%d.%m.%Y")
		for i in range(self.income_json['recordsTotal']):
			now = datetime.datetime.strptime(self.income_json['transactions'][i]['date'], "%d.%m.%Y")
			if (self.income_json['transactions'][i]['project']['id'] == id_proj) and (start <= now <= end):
				income += float(self.income_json['transactions'][i]['amount'])
		return income


class Accountant:
	def calculating_wages(self, worked_hours, standard_hours, salary_rate):
		return (worked_hours/standard_hours)*salary_rate
