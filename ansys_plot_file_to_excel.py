import matplotlib.pyplot as plt
import tkinter
import xlwt
from functools import partial
from tkinter import font as tkfont
from tkinter.filedialog import askopenfilenames, asksaveasfilename
from typing import Dict, List, Optional, Tuple, Union

ico_open = u"\U0001F4C2"
ico_save = u"\U0001F4BE"
ico_exit = u"\U0001F6AA"
ico_chart = u"\U0001F4C8"


def obtain_data(file_name: str) -> Dict[str, Union[List[str], str, float]]:
    with open(f'{file_name}', 'r') as file:
        data: List[str] = file.readlines()
        data_out: List[str] = []
        for index, line in enumerate(data):
            if index > 4:
                if line.split('\t')[0] == '0.06':
                    data_out.append(line)
                    break
        for index, line in enumerate(data):
            if index > 4:
                if line.split('\t')[0] != '-0.06':
                    if float(line.split('\t')[0]) < \
                            float(data_out[-1].split('\t')[0]):
                        data_out.append(line)
                else:
                    data_out.append(line)
                    break
        if 'temp' in file_name.lower():
            correction = 273.15
        else:
            correction = 0
        return {
            'data_out': data_out,
            'coord_name': data[1].split('"')[1].strip('"'),
            'value_name': data[1].split('"')[3].strip('"'),
            'correction': correction
        }


class App:

    def __init__(self) -> None:
        self.root = tkinter.Tk(className='ansys to excel')
        self.font = tkfont.Font(family='Arial', size=14, weight='bold')
        self.root.rowconfigure((0, 1), weight=1)
        self.root.columnconfigure((0, 1), weight=1)
        tkinter.Button(
            self.root,
            text=f'{ico_open} Open files',
            command=self.open_files,
            font=self.font).grid(
            row=0, column=0, columnspan=1, padx=10, pady=20, sticky='EWNS'
        )
        tkinter.Button(
            self.root,
            text=f'{ico_exit} Quit',
            command=self.quit, font=self.font
        ).grid(
            row=0, column=1, columnspan=1, padx=10, pady=20, sticky='EWNS'
        )
        self.file_names: Optional[Tuple[str]] = None
        self.plot_buttons: List[tkinter.Button] = []
        self.write_excel_button: Optional[tkinter.Button] = None
        self.current_plot: Optional[str] = None
        self.root.mainloop()

    def open_files(self) -> None:
        self.file_names = askopenfilenames()
        if self.plot_buttons or self.write_excel_button:
            for button in self.plot_buttons:
                button.grid_forget()
            self.plot_buttons = []
            self.write_excel_button.grid_forget()
            self.write_excel_button = None

        self.make_buttons_for_files()

    def write_to_excel(self):
        workbook = xlwt.Workbook()
        for file_name in self.file_names:
            file_data: Dict = obtain_data(file_name)
            worksheet = workbook.add_sheet(
                f'{file_name.split("/")[-1]}_out')
            worksheet.write(0, 0, file_data['coord_name'])
            worksheet.write(0, 1, file_data['value_name'])
            for index, line in enumerate(file_data['data_out']):
                worksheet.write(
                    index + 2, 0,
                    float(line.split('\t')[0]) * 1000 + 60
                )
                worksheet.write(
                    index + 2, 1,
                    float(line.split('\t')[1]) - file_data['correction']
                )
        out_file_name: str = asksaveasfilename(
            defaultextension=".xls", filetypes=(
                ("Microsoft Excel file", "*.xls"),
                ("All Files", "*.*")
            ))
        try:
            workbook.save(f'{out_file_name}')
        except FileNotFoundError:
            pass

    def make_buttons_for_files(self):
        row = 0
        for index, file_name in enumerate(self.file_names):
            if index % 2 == 0:
                row += 1
            button = tkinter.Button(
                self.root,
                text=f'{ico_chart} Plot {file_name.split("/")[-1]}',
                command=partial(self.plot_file, file_name),
                font=self.font
            )
            self.plot_buttons.append(button)
            button.grid(
                row=row, column=index % 2, columnspan=1,
                padx=10, pady=5,
                sticky='EWNS'
            )

        self.write_excel_button = tkinter.Button(
            self.root,
            text=f'{ico_save} Save all to xls file',
            command=self.write_to_excel, font=self.font
        )
        self.write_excel_button.grid(
            row=row + 1, column=0, columnspan=2,
            padx=10, pady=20, sticky='EWNS'
        )
        self.root.rowconfigure((0, row), weight=1)

    def plot_file(self, file_name: str) -> None:
        plot_data: Dict = obtain_data(file_name)
        plt.title(
            f'{file_name.split("/")[-1].replace("_", " ")} = f(coordinate)')
        plt.xlabel(f'Coordinate, mm')
        if plot_data['correction'] == 0:
            plt.ylabel('Velocity, m/s')
            if not self.current_plot:
                self.current_plot = 'velocity'
            elif self.current_plot == 'temperature':
                plt.close()
                self.current_plot = 'velocity'
        else:
            plt.ylabel('Temperature, oC')
            if not self.current_plot:
                self.current_plot = 'temperature'
            elif self.current_plot == 'velocity':
                plt.close()
                self.current_plot = 'temperature'
        plt.grid(visible=True)
        plt.minorticks_on()
        x_axis_data: List[float] = []
        y_axis_data: List[float] = []
        for line in plot_data['data_out']:
            x_axis_data.append(
                float(line.split('\t')[0]) * 1000 + 60
            )
            y_axis_data.append(
                float(line.split('\t')[1]) - plot_data['correction']
            )
        plt.plot(
            x_axis_data, y_axis_data,
            label=f'{file_name.split("/")[-1].replace("_", " ")}'
        )
        plt.legend()

        plt.show()

    def quit(self):
        plt.close()
        self.root.destroy()


app = App()
