import time
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import os


class Window:
	def __init__(self):
		self.api_text_path = os.path.abspath(rf'C:\Users\{os.getlogin()}\.accountant\api.txt')
		self.api_text = ''
		self.flag_ready = False
		self.PAD = 3
		self.WIDTH = 40
		self.months = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Ноябрь', 'Октябрь', 'Декабрь']
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
		self.root.bind_all("<Key>", self._onKeyRelease, "+")

		self.frm = tk.Frame(self.root)
		self.frm.pack()

		tk.Label(self.frm, text='Adesk API').grid(column=0, row=0, sticky='w', padx=self.PAD, pady=self.PAD)
		tk.Label(self.frm, text='Месяц').grid(column=0, row=1, sticky='w', padx=self.PAD, pady=self.PAD)
		tk.Label(self.frm, text='Путь к Excel').grid(column=0, row=2, sticky='w', padx=self.PAD, pady=self.PAD)
		tk.Label(self.frm, text='Путь к Bitrix').grid(column=0, row=3, sticky='w', padx=self.PAD, pady=self.PAD)

		self.api = tk.Entry(self.frm, width=self.WIDTH)
		self.api.insert(0, self.api_text)
		self.month = ttk.Combobox(self.frm, width=self.WIDTH-3, values=self.months)
		self.excel_path = tk.Entry(self.frm, width=self.WIDTH)
		self.bitrix_path = tk.Entry(self.frm, width=self.WIDTH)
		self.api.grid(column=1, row=0, padx=self.PAD, pady=self.PAD)
		self.month.grid(column=1, row=1, padx=self.PAD, pady=self.PAD)
		self.excel_path.grid(column=1, row=2, padx=self.PAD, pady=self.PAD)
		self.bitrix_path.grid(column=1, row=3, padx=self.PAD, pady=self.PAD)

		self.str_btn = tk.Button(self.frm, text='Запуск', command=self.ready)
		self.str_btn.grid(column=0, row=4, columnspan=2, padx=self.PAD, pady=self.PAD)

	def _onKeyRelease(self, event):
		ctrl = (event.state & 0x4) != 0
		if event.keycode == 88 and ctrl and event.keysym.lower() != "x":
			event.widget.event_generate("<<Cut>>")

		if event.keycode == 86 and ctrl and event.keysym.lower() != "v":
			event.widget.event_generate("<<Paste>>")

		if event.keycode == 67 and ctrl and event.keysym.lower() != "c":
			event.widget.event_generate("<<Copy>>")

	def get_data(self):
		if self.api_text != self.api.get():
			with open(self.api_text_path, 'w', encoding='utf-8') as f:
				f.write(self.api.get())
		data = (self.api.get(), self.month.get(), self.excel_path.get().strip(), self.bitrix_path.get().strip())
		self.flag_ready = False
		if '' in data:
			messagebox.showerror(message='Вы не указали один из параметров')
			return False
		else:
			return tuple(data)

	def ready(self):
		self.flag_ready = True

	def get_response_ready(self):
		while not self.flag_ready:
			time.sleep(0.1)

	def mainloop(self):
		self.root.mainloop()
