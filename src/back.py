import bs4
import requests
import openpyxl
import datetime


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
			# for col in range(1, self.worksheet.max_column):
			self.data[self.worksheet[row][1].value] = [self.worksheet[row][2].value, [dict([(int(self.worksheet[2][i].value[:4]), self.worksheet[row][i].value)]) for i in range(3, self.worksheet.max_row-3)]]
		return self.data
