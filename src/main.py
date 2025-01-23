import time
import back
import front
import threading as th


def main():
	while True:
		window.get_response_ready()
		data = window.get_data()
		adesk = back.Adesk(data[0])
		consultant = back.Consultant()
		# excel = back.Excel(data[2])
		bitrix = back.Bitrix(data[3])
		print(bitrix.get_projects_in_month())


if __name__ == '__main__':
	consul = back.Consultant()
	window = front.Window()
	th.Thread(target=main, daemon=True).start()
	window.mainloop()

# excel = back.Excel(r'D:\Dima\Python\ABE\beginning.xlsx')
# print(excel.get_jobs_and_div())
# bit = back.Bitrix(r'D:\Dima\Python\ABE\Bitrix24_hours.xlsx')
# print(bit.get_hours())
