import time
import back
import front
import threading as th


def main():
	window.get_response_ready()
	data = window.get_data()
	print(data)


if __name__ == '__main__':
	window = front.Window()
	th.Thread(target=main, daemon=True).start()
	window.mainloop()

# excel = back.Excel(r'D:\Dima\Python\ABE\beginning.xlsx')
# print(excel.get_jobs_and_div())
# bit = back.Bitrix(r'D:\Dima\Python\ABE\Bitrix24_hours.xlsx')
# print(bit.get_hours())
