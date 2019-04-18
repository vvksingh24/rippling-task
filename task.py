import csv
import os
import xlrd


cwd = os.getcwd()
csv_path = os.path.join(cwd, "Csv's")
if not os.path.exists(csv_path):
	os.makedirs(csv_path)

try:
	input_file = input("enter the path of the Excel sheet: ")
	xcel_file = xlrd.open_workbook(input_file)
	for sheet in xcel_file.sheet_names():
		xl_sheet = xcel_file.sheet_by_name(sheet)
		heading = xl_sheet.row_values(0)
		for row_index in range(1,xl_sheet.nrows):
			line = xl_sheet.row_values(row_index)
			employee_file_path = os.path.join(csv_path, line[0].split("@")[0] + ".csv")
			if os.path.exists(employee_file_path):
				with open(employee_file_path, 'a') as csv_file:
					csvwriter = csv.writer(csv_file)
					csvwriter.writerow(line)
			else:
				with open(employee_file_path, 'a') as csv_file:
					csvwriter = csv.writer(csv_file)
					csvwriter.writerow(heading)
					csvwriter.writerow(line)

except IOError:
	print ("file is not present in the location")
except Exception as e:
	raise(e)