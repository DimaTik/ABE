import tkinter as tk
from tkinter import ttk


class Window:
	def __init__(self):
		self.PAD = 3
		self.WIDTH = 40
		self.root = tk.Tk()
		self.root.title('Подсчет ЗП')
		# self.root.geometry('500x500')

		self.frm = tk.Frame(self.root)
		self.frm.pack()

		tk.Label(self.frm, text='Adesk API').grid(column=0, row=0, sticky='w', padx=self.PAD, pady=self.PAD)
		tk.Label(self.frm, text='Месяц').grid(column=0, row=1, sticky='w', padx=self.PAD, pady=self.PAD)
		tk.Label(self.frm, text='Путь к Excel').grid(column=0, row=2, sticky='w', padx=self.PAD, pady=self.PAD)
		tk.Label(self.frm, text='Путь к Bitrix').grid(column=0, row=3, sticky='w', padx=self.PAD, pady=self.PAD)

		self.api = tk.Entry(self.frm, width=self.WIDTH)
		self.month = ttk.Combobox(self.frm, width=self.WIDTH-3)
		self.excel_path = tk.Entry(self.frm, width=self.WIDTH)
		self.bitrix_path = tk.Entry(self.frm, width=self.WIDTH)
		self.api.grid(column=1, row=0, padx=self.PAD, pady=self.PAD)
		self.month.grid(column=1, row=1, padx=self.PAD, pady=self.PAD)
		self.excel_path.grid(column=1, row=2, padx=self.PAD, pady=self.PAD)
		self.bitrix_path.grid(column=1, row=3, padx=self.PAD, pady=self.PAD)

		self.str_btn = tk.Button(self.frm, text='Запуск')
		self.str_btn.grid(column=0, row=4, columnspan=2, padx=self.PAD, pady=self.PAD)

		self.root.mainloop()

	def get_data(self):
		pass
