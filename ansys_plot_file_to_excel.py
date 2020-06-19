from tkinter.filedialog import askopenfilename

import matplotlib.pyplot as plt

import xlwt

file_name = askopenfilename()


def plot(a, b):
    plt.title(f'{file_name.split("_")[-1]} = f(coordinate)')
    plt.xlabel(f'{file_name.split("_")[-1]}')
    plt.ylabel('temperature')
    # plt.legend('x = y^2; x = y')
    plt.grid()
    plt.minorticks_on()
    plt.plot(a, b)
    plt.show()


with open(f'{file_name}', 'r') as file:
    data = file.readlines()
    data_out = []
    for i, line in enumerate(data):
        if i > 4:
            if line.split('\t')[0] == '0.06':
                data_out.append(line)
                break

    for i, line in enumerate(data):
        if i > 4:
            if line.split('\t')[0] != '-0.06':
                if float(line.split('\t')[0]) < \
                        float(data_out[-1].split('\t')[0]):
                    data_out.append(line)
            else:
                data_out.append(line)
                break

workbook = xlwt.Workbook()
worksheet = workbook.add_sheet(f'{file_name.split("/")[-1]}_out')

# write output data to xls
worksheet.write(0, 0, data[1].split('"')[1].strip('"'))
worksheet.write(0, 1, data[1].split('"')[3].strip('"'))
x_axis_data = []
y_axis_data = []
for i, line in enumerate(data_out):
    worksheet.write(i + 2, 0, float(line.split('\t')[0]))
    x_axis_data.append(float(line.split('\t')[0]))
    worksheet.write(i + 2, 1, float(line.split('\t')[1]) - 273.15)
    y_axis_data.append(float(line.split('\t')[1]) - 273.15)

# write raw data to second sheet of xls
worksheet2 = workbook.add_sheet(f'{file_name.split("/")[-1]}_raw')
worksheet2.write(0, 0, data[1].split('"')[1].strip('"'))
worksheet2.write(0, 1, data[1].split('"')[3].strip('"'))
for i, line in enumerate(data):
    if 4 < i < len(data) - 1:
        worksheet2.write(i - 3, 0, float(line.split('\t')[0]))
        worksheet2.write(i - 3, 1, float(line.split('\t')[1]) - 273.15)

# write txt file with output data
with open(f'{file_name}_out.txt', 'w') as out_file:
    for line in data_out:
        out_file.write(line)

# saving xls file
workbook.save(f'{file_name}_out.xls')

# plotting output data
plot(x_axis_data, y_axis_data)
