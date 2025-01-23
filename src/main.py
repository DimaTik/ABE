import time
import back
import front
import threading as th


def main():
	window.get_response_ready()
	data = window.get_data()
	# adesk = back.Adesk(data[0])
	# consultant = back.Consultant(data[1])
	# excel = back.Excel(data[2])
	# bitrix = back.Bitrix(data[3])


if __name__ == '__main__':
	consul = back.Consultant()
	print(consul.get_standard_hours('Январь'))
	window = front.Window()
	th.Thread(target=main, daemon=True).start()
	window.mainloop()

# excel = back.Excel(r'D:\Dima\Python\ABE\beginning.xlsx')
# print(excel.get_jobs_and_div())
# bit = back.Bitrix(r'D:\Dima\Python\ABE\Bitrix24_hours.xlsx')
# print(bit.get_hours())
