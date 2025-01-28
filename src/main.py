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
			excel = back.Excel(data_from_user[1], data_from_user[2])
			bitrix = back.Bitrix(data_from_user[3])
			account = back.Accountant()

			projects = adesk.get_projects() 	# Нет доступа
			# projects = ['79', '81', '52']
			worked_hours = bitrix.get_hours_worked()
			standard_hours = consultant.get_standard_hours(data_from_user[1])
			data_of_workers = excel.get_jobs_and_div()

			print(worked_hours)
			print(standard_hours)
			print(data_of_workers)

			for project in projects:
				res_data = [project, []]
				for i in range(len(worked_hours[project])):
					name = list(worked_hours[project])[i]
					hours = list(worked_hours[project].values())[i]
					salary_rate = data_of_workers[list(worked_hours[project].keys())[i]][-1]
					job = data_of_workers[list(worked_hours[project])[i]][0]
					div = data_of_workers[list(worked_hours[project])[i]][1]
					res_data[1].append([name, job, div, hours, account.calculating_wages(hours, standard_hours, salary_rate)])
				excel.create_result_table(res_data)
				excel.set_result_table(res_data)
			excel.del_sheet1()


if __name__ == '__main__':
	consul = back.Consultant()
	window = front.Window()
	th.Thread(target=main, daemon=True).start()
	window.mainloop()
