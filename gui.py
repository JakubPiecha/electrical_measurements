import customtkinter as ctk
import os
from collections import defaultdict
from openpyxl import load_workbook
from tkcalendar import DateEntry

ctk.set_appearance_mode('dark')

ctk.set_default_color_theme('green')


class MyFrame(ctk.CTkScrollableFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.params = {'padx': 10, 'pady': 15, 'sticky': 'w'}

        self.data = ElectricalMeasurementsGenerate()

        self.front_page = ctk.CTkCheckBox(self, text='Strona tytułowa', onvalue='front_page', offvalue=0)
        self.front_page.grid(row=0, column=0, **self.params)

        self.zeroing = ctk.CTkCheckBox(self, text='Zerowanie', onvalue='zeroing', offvalue=0)
        self.zeroing.grid(row=0, column=1, **self.params)

        self.differential = ctk.CTkCheckBox(self, text='Różnicówka', onvalue='differential', offvalue=0)
        self.differential.grid(row=0, column=2, **self.params)

        self.insulation = ctk.CTkCheckBox(self, text='Izolacje', onvalue='insulation', offvalue=0)
        self.insulation.grid(row=0, column=3, **self.params)

        self.reflexes = ctk.CTkCheckBox(self, text='Odgromy', onvalue='reflexes', offvalue=0)
        self.reflexes.grid(row=0, column=4, **self.params)

        self.investment_data_label = ctk.CTkLabel(self, text='Dane obiektu wykonanych pomiarów:')
        self.investment_data_label.grid(row=1, column=0, **self.params)

        self.investment_zip_code_entry = ctk.CTkEntry(self, placeholder_text='Kod pocztowy')
        self.investment_zip_code_entry.grid(row=1, column=1, **self.params)

        self.investment_city_entry = ctk.CTkEntry(self, placeholder_text='Miejscowość')
        self.investment_city_entry.grid(row=1, column=2, **self.params)

        self.investment_street_and_number_entry = ctk.CTkEntry(self, width=180,
                                                               placeholder_text='ulica, nr domu/mieszkania')
        self.investment_street_and_number_entry.grid(row=1, column=3, **self.params)

        self.investment_name_entry = ctk.CTkEntry(self, placeholder_text='Nazwa Obiektu')
        self.investment_name_entry.grid(row=1, column=4, **self.params)

        self.investors_label = ctk.CTkLabel(self, text='Dane Zleceniodawcy pomiarów:')
        self.investors_label.grid(row=2, column=0, **self.params)

        self.investors_zip_code_entry = ctk.CTkEntry(self, placeholder_text='Kod pocztowy')
        self.investors_zip_code_entry.grid(row=2, column=1, **self.params)

        self.investors_city_entry = ctk.CTkEntry(self, placeholder_text='Miejscowość')
        self.investors_city_entry.grid(row=2, column=2, **self.params)

        self.investors_street_and_number_entry = ctk.CTkEntry(self, width=180,
                                                              placeholder_text='ulica, nr domu/mieszkania')
        self.investors_street_and_number_entry.grid(row=2, column=3, **self.params)

        self.investors_name_entry = ctk.CTkEntry(self, placeholder_text='Nazwa')
        self.investors_name_entry.grid(row=2, column=4, **self.params)

        self.date_label = ctk.CTkLabel(self, text='Data wykonania pomiarów:')
        self.date_label.grid(row=3, column=0, **self.params)

        self.date = DateEntry(self, selectmode='day', background='dark', foreground='grey',
                              selectbackground='grey', locale='pl')
        self.date.grid(row=3, column=1, **self.params)

        self.storey_name_label = ctk.CTkLabel(self, text='Nazwa kondygnacji')
        self.storey_name_label.grid(row=4, column=0, **self.params)

        self.storey_name = ctk.CTkEntry(self, placeholder_text='np. Parter')
        self.storey_name.grid(row=4, column=1, **self.params)

        self.add_storey_button = ctk.CTkButton(self, text='Dodaj kondygnację',
                                               command=lambda: self.add_storey(self.storey_name.get()))
        self.add_storey_button.grid(row=5, column=1, **self.params)

        self.add = ctk.CTkCheckBox(self, text='dodaj', command=self.add_storey,  onvalue='front_page', offvalue=0)
        self.add.grid(row=6, column=0, **self.params)

    def add_storey(self, storey_name='próba'):
        print(self.add.get())
        last_row = self.grid_size()[-1]
        self.data.order_data['floors'][storey_name] = defaultdict(dict)
        self.storey_name.delete(0, 'end')

        room_name_Label = ctk.CTkLabel(self, text=f'Nazwa/Liczba gniazdek dla - {storey_name}')
        room_name_Label.grid(row=last_row + 1, column=0, **self.params)

        room_name = ctk.CTkEntry(self, placeholder_text='Nazwa pomieszczenia')
        room_name.grid(row=last_row + 1, column=1, **self.params)

        power_sockets = ctk.CTkEntry(self, placeholder_text='Liczba gniazdek')
        power_sockets.grid(row=last_row + 1, column=2, **self.params)

        add_storey_button = ctk.CTkButton(self, text='Dodaj',
                                          command=lambda: self.add_room(storey_name, room_name, power_sockets))
        add_storey_button.grid(row=last_row + 1, column=3, **self.params)

    def add_room(self, storey_name, room_name, power_sockets):
        self.data.order_data['floors'][storey_name].update({room_name.get(): power_sockets.get()})
        room_name.delete(0, 'end')
        power_sockets.delete(0, 'end')


class App(ctk.CTk):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.title('Electrical Measurements')

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.my_frame = MyFrame(master=self, width=1000, height=600, corner_radius=0, fg_color="transparent")
        self.my_frame.grid(row=0, column=0, sticky="nsew")

        self.generate_button = ctk.CTkButton(self, text='Generuj pliki', command=self.generate)
        self.generate_button.grid(row=1, column=0, padx=10, pady=15, sticky='e')

    def generate(self):
        obj = self.my_frame.data.order_data
        obj['investment_zip_code'] = self.my_frame.investment_zip_code_entry.get()
        obj['investment_name'] = self.my_frame.investment_name_entry.get()
        obj['investment_street_and_number'] = self.my_frame.investment_street_and_number_entry.get()
        obj['investment_city'] = self.my_frame.investment_city_entry.get()
        obj['investors_zip_code'] = self.my_frame.investors_zip_code_entry.get()
        obj['investors_name'] = self.my_frame.investors_name_entry.get()
        obj['investors_street_and_number'] = self.my_frame.investors_street_and_number_entry.get()
        obj['investors_city'] = self.my_frame.investors_city_entry.get()
        obj['date'] = self.my_frame.date.get()

        self.my_frame.data.create_directory()

        call_dict = {
            self.my_frame.front_page.get(): self.my_frame.data.print_front_page,
            self.my_frame.differential.get(): self.my_frame.data.differential,
            self.my_frame.zeroing.get(): self.my_frame.data.zeroing,
            self.my_frame.insulation.get(): self.my_frame.data.insulation,
            self.my_frame.reflexes.get(): self.my_frame.data.reflexes
        }

        for key in call_dict.keys():
            if key:
                call_dict[key]()


class ElectricalMeasurementsGenerate:
    def __init__(self):
        self.order_data = {
            'investment_zip_code': '',
            'investment_name': '',
            'investment_street_and_number': '',
            'investment_city': '',
            'investors_zip_code': '',
            'investors_name': '',
            'investors_street_and_number': '',
            'investors_city': '',
            'date': '',
            'floors': defaultdict()

        }
        self.path = ''

    def create_directory(self):
        investors_name = self.remove_special_char("_".join(self.order_data['investors_name'].split()))
        objects_name = self.remove_special_char("_".join(self.order_data['investment_name'].split()))
        objects_adres = self.remove_special_char("_".join(self.order_data['investment_street_and_number'].split()))

        directory = f'{investors_name}_{objects_name}_{objects_adres}'
        self.path = os.path.join(os.path.abspath('..'), directory)
        os.mkdir(self.path)

    def save_file(self, filename, file):
        savefile = f'{filename.split(".")[0]}_{"_".join((self.order_data["investment_name"]).split())}.xlsx'
        file.save(os.path.join(self.path, savefile))

    def insert_column(self, wb):
        row = 10
        ws = wb.active

        ws['A7'] = f'{self.order_data["investment_zip_code"]} {self.order_data["investment_city"]},' \
                   f' {self.order_data["investment_street_and_number"]} - {self.order_data["investment_name"]}'

        for name, floor in self.order_data['floors'].items():
            ws[f'C{str(row)}'] = name
            row += 1

            for room, number in floor.items():
                ws[f'C{str(row)}'] = room
                row += 1

                for ps in range(1, int(number) + 1):
                    ws[f'C{str(row)}'] = f'Gniazdo jednofazowe A/Z 230V nr {ps}'
                    ws[f'B{str(row)}'] = ps
                    row += 1
        return wb

    def add_to_register(self):
        filename = 'rejestr.xlsx'
        wb = load_workbook(filename=filename)
        ws = wb.active
        max_row = wb.active.max_row
        last_year = ws[f'C{max_row}'].value.split('.')[-1]
        year = self.order_data['date'].split('.')[-1]

        ws[f'A{max_row + 1}'] = f'{self.order_data["investors_name"]} - {self.order_data["investment_name"]} - ' \
                                f'{self.order_data["investment_street_and_number"]}'
        number = ws[f'B{max_row + 1}'] = ws[f'B{max_row}'].value + 1 if last_year == year else 1
        ws[f'C{max_row + 1}'] = self.order_data['date']

        wb.save('rejestr.xlsx')

        return f'{number}/{year}'

    def print_front_page(self):
        filename = 'strona_tytulowa.xlsx'
        wb = load_workbook(filename=filename)
        ws = wb.active
        ws['B6'] = self.order_data['investors_name']
        ws['B7'] = self.order_data['investors_street_and_number']
        ws['B8'] = self.order_data['investors_zip_code'] + ' ' + self.order_data['investment_city']
        ws['B9'] = self.order_data['investment_name']
        ws['B10'] = self.order_data['investment_street_and_number']
        ws['B11'] = self.order_data['investment_zip_code'] + ' ' + self.order_data['investment_city']
        ws['B17'] = self.order_data['date']
        date = self.order_data['date'].split('.')
        next_date = date[1] + "/" + (str(int(date[-1]) + 5))
        ws['B28'] = next_date
        ws['B2'] = self.add_to_register()

        self.save_file(filename, wb)

    def zeroing(self):
        filename = 'zerowanie.xlsx'
        wb = load_workbook(filename=filename)
        self.insert_column(wb)

        self.save_file(filename, wb)

    def differential(self):
        filename = 'roznicowka.xlsx'
        wb = load_workbook(filename=filename)
        self.insert_column(wb)
        self.save_file(filename, wb)

    def insulation(self):
        filename = 'izo.xlsx'
        wb = load_workbook(filename=filename)
        ws = wb.active

        ws['A7'] = f'{self.order_data["investment_zip_code"]} {self.order_data["investment_city"]},' \
                   f' {self.order_data["investment_street_and_number"]} - {self.order_data["investment_name"]}'

    def reflexes(self):
        filename = 'odgromy.xlsx'
        wb = load_workbook(filename=filename)
        ws = wb.active

        ws['A7'] = f'{self.order_data["investment_zip_code"]} {self.order_data["investment_city"]},' \
                   f' {self.order_data["investment_street_and_number"]} - {self.order_data["investment_name"]}'

    @staticmethod
    def remove_special_char(string):
        special_char = ['\\', '/', '[', ']', ':', '¦', '<', '>', '+', '=', ';', ',', '*', '?', '"']
        new_string = "".join(filter(lambda char: char not in special_char, string))
        return new_string


app = App()

app.mainloop()
