import time
import back
import front
import threading as th


def main():
	while True:
		window.get_response_ready()
		data_from_user = window.get_data()
		if data_from_user:
			adesk = back.Adesk(data_from_user[0])
			consultant = back.Consultant()
			excel = back.Excel(data_from_user[2])
			bitrix = back.Bitrix(data_from_user[3])
			account = back.Accountant()

			projects = bitrix.get_projects_in_month()
			# month_from_bit = bitrix.get_month_from_name()
			hours_worked = bitrix.get_hours_worked()
			standard_hours = consultant.get_standard_hours(data_from_user[1])
			data_of_workers = excel.get_jobs_and_div()

			print(hours_worked)
			print(standard_hours)
			print(data_of_workers)

			for project in projects:
				pass
				# excel.create_result_table(...)
				# excel.set_result_table(...)
			# excel.create_result_table()
			# excel.create_result_table()


if __name__ == '__main__':
	consul = back.Consultant()
	window = front.Window()
	th.Thread(target=main, daemon=True).start()
	window.mainloop()

# excel = back.Excel(r'D:\Dima\Python\ABE\beginning.xlsx')
# print(excel.get_jobs_and_div())
# bit = back.Bitrix(r'D:\Dima\Python\ABE\Bitrix24_hours.xlsx')
# print(bit.get_hours())
