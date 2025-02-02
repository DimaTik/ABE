import time
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import os


class Window:
	def __init__(self):
		self.api_text_path = os.path.abspath(rf'C:\Users\{os.getlogin()}\.accountant\api.txt')
		self.api_text = ''
		self.flag_ready = False
		self.PAD = 3
		self.WIDTH = 40
		self.months = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Ноябрь', 'Октябрь', 'Декабрь']
		self.years = [str(i) for i in range(2000, 2051)]
		if os.path.isfile(self.api_text_path):
			with open(self.api_text_path, encoding='utf-8') as f:
				self.api_text = f.readline()
		else:
			os.mkdir(rf'C:\Users\{os.getlogin()}\.accountant')
			with open(self.api_text_path, 'w', encoding='utf-8') as f:
				f.write('0')

		self.root = tk.Tk()
		self.root.title('Accountant')
		self.root.resizable(False, False)
		self.root.bind_all("<Key>", self.__onKeyRelease, "+")

		self.frm = tk.Frame(self.root)
		self.frm.pack()

		for row, text in enumerate(('Adesk API', 'Месяц', 'Год', 'Путь к Excel', 'Путь к Bitrix')):
			tk.Label(self.frm, text=text).grid(column=0, row=row, sticky='w', padx=self.PAD, pady=self.PAD)

		for row, command in enumerate((self.__get_path_to_excel, self.__get_directory_to_bitrix), 3):
			tk.Button(self.frm, text='Открыть', command=command)\
				.grid(column=3, row=row, sticky='w', padx=self.PAD, pady=self.PAD)

		self.api = tk.Entry(self.frm, width=self.WIDTH)
		self.api.insert(0, self.api_text)
		self.month = ttk.Combobox(self.frm, width=self.WIDTH-3, values=self.months)
		self.year = ttk.Combobox(self.frm, width=self.WIDTH-3, values=self.years)
		self.excel_path = tk.Entry(self.frm, width=self.WIDTH)
		self.bitrix_directory = tk.Entry(self.frm, width=self.WIDTH)

		for row, object in enumerate((self.api, self.month, self.year, self.excel_path, self.bitrix_directory)):
			object.grid(column=1, row=row, padx=self.PAD, pady=self.PAD)

		self.str_btn = tk.Button(self.frm, text='Запуск', command=self.__ready)
		self.str_btn.grid(column=0, row=5, columnspan=2, padx=self.PAD, pady=self.PAD)

	def __onKeyRelease(self, event):
		ctrl = (event.state & 0x4) != 0
		if event.keycode == 88 and ctrl and event.keysym.lower() != "x":
			event.widget.event_generate("<<Cut>>")

		if event.keycode == 86 and ctrl and event.keysym.lower() != "v":
			event.widget.event_generate("<<Paste>>")

		if event.keycode == 67 and ctrl and event.keysym.lower() != "c":
			event.widget.event_generate("<<Copy>>")

	def __get_path_to_excel(self):
		path = filedialog.askopenfilename()
		if path != '':
			self.excel_path.delete(0, tk.END)
			self.excel_path.insert(0, path)

	def __get_directory_to_bitrix(self):
		directory = filedialog.askdirectory()
		if directory != '':
			self.bitrix_directory.delete(0, tk.END)
			self.bitrix_directory.insert(0, directory)

	def get_data(self):
		if self.api_text != self.api.get():
			with open(self.api_text_path, 'w', encoding='utf-8') as f:
				f.write(self.api.get())
		data = (self.api.get(), self.month.get(), self.year.get(), self.excel_path.get().strip(), self.bitrix_directory.get().strip())
		self.flag_ready = False
		if '' in data:
			messagebox.showerror(message='Вы не указали один из параметров')
			return None
		if not os.path.exists(fr'{self.bitrix_directory.get().strip()}\{self.month.get()}_{self.year.get()}_bitrix.xlsx'):
			messagebox.showerror(message='Не найден файл Bitrix')
			return None
		else:
			return data

	def __ready(self):
		self.flag_ready = True

	def get_response_ready(self):
		while not self.flag_ready:
			time.sleep(0.1)

	def mainloop(self):
		self.root.mainloop()
