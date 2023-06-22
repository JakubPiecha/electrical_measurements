import customtkinter as ctk
from CTkMessagebox import CTkMessagebox
from tkcalendar import DateEntry

from file_generate.electrical_measurements_generate import ElectricalMeasurementsGenerate
from settings_window.settings_window import SettingsWindow
from frames.zeroing_frame import ZeroingFrame
from frames.differential_frame import DifferentialFrame
from frames.reflexes_frame import ReflexesFrame


class MainFrame(ctk.CTkScrollableFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        self.generator_object = ElectricalMeasurementsGenerate()

        self.params = {'padx': 8, 'pady': 8, 'sticky': 'w'}

        self.settings_button = ctk.CTkButton(self, text="Ustawienia", command=self.open_settings_window)
        self.settings_button.grid(row=0, column=0, **self.params)

        self.settings_window = None

        self.front_page = ctk.CTkCheckBox(self, text='Strona tytułowa', onvalue='front_page', offvalue=0)
        self.front_page.grid(row=1, column=0, **self.params)

        self.zeroing = ctk.CTkCheckBox(self,
                                       text='Zerowanie', onvalue='zeroing', offvalue=0, command=self.show_zeroing)
        self.zeroing.grid(row=1, column=1, **self.params)

        self.differential = ctk.CTkCheckBox(self, text='Różnicówka', onvalue='differential', offvalue=0,
                                            command=self.show_differential)
        self.differential.grid(row=1, column=2, **self.params)

        self.insulation = ctk.CTkCheckBox(self, text='Izolacje', onvalue='insulation', offvalue=0)
        self.insulation.grid(row=1, column=3, **self.params)

        self.reflexes = ctk.CTkCheckBox(self, text='Odgromy', onvalue='reflexes', offvalue=0,
                                        command=self.show_reflexes)
        self.reflexes.grid(row=1, column=4, **self.params)

        self.investment_data_label = ctk.CTkLabel(self, text='Dane obiektu wykonanych pomiarów:')
        self.investment_data_label.grid(row=2, column=0, **self.params)

        self.investment_zip_code_entry = ctk.CTkEntry(self, placeholder_text='Kod pocztowy')
        self.investment_zip_code_entry.grid(row=2, column=1, **self.params)

        self.investment_city_entry = ctk.CTkEntry(self, placeholder_text='Miejscowość')
        self.investment_city_entry.grid(row=2, column=2, **self.params)

        self.investment_street_and_number_entry = ctk.CTkEntry(self, width=180,
                                                               placeholder_text='ulica, nr domu/mieszkania')
        self.investment_street_and_number_entry.grid(row=2, column=3, **self.params)

        self.investment_name_entry = ctk.CTkEntry(self, placeholder_text='Nazwa Obiektu')
        self.investment_name_entry.grid(row=2, column=4, **self.params)

        self.investors_label = ctk.CTkLabel(self, text='Dane Zleceniodawcy pomiarów:')
        self.investors_label.grid(row=3, column=0, **self.params)

        self.investors_zip_code_entry = ctk.CTkEntry(self, placeholder_text='Kod pocztowy')
        self.investors_zip_code_entry.grid(row=3, column=1, **self.params)

        self.investors_city_entry = ctk.CTkEntry(self, placeholder_text='Miejscowość')
        self.investors_city_entry.grid(row=3, column=2, **self.params)

        self.investors_street_and_number_entry = ctk.CTkEntry(self, width=180,
                                                              placeholder_text='ulica, nr domu/mieszkania')
        self.investors_street_and_number_entry.grid(row=3, column=3, **self.params)

        self.investors_name_entry = ctk.CTkEntry(self, placeholder_text='Nazwa')
        self.investors_name_entry.grid(row=3, column=4, **self.params)

        self.date_label = ctk.CTkLabel(self, text='Data wykonania pomiarów:')
        self.date_label.grid(row=4, column=0, **self.params)

        self.date = DateEntry(self, selectmode='day', background='dark', foreground='grey',
                              selectbackground='grey', locale='pl')
        self.date.grid(row=4, column=1, **self.params)

        self.generate_button = ctk.CTkButton(self, text='Generuj pliki', command=self.generate)
        self.generate_button.grid(row=9, column=5, padx=10, pady=15, sticky='w')

        self.zeroing_frame = ZeroingFrame(master=self, height=200, width=1150, corner_radius=10, fg_color="gray20",
                                          label_text='Zerowanie', label_anchor='center')

        self.differential_frame = DifferentialFrame(master=self, width=1150, height=200, corner_radius=10,
                                                    fg_color="gray20", label_text='Różnicówka', label_anchor='center')

        self.reflexes_frame = ReflexesFrame(master=self, corner_radius=10, fg_color="gray20")

    def open_settings_window(self):
        if self.settings_window is None or not self.settings_window.winfo_exists():
            self.settings_window = SettingsWindow(self)
            self.settings_window.grab_set()
        else:
            self.settings_window.focus()

    def show_zeroing(self):
        if self.zeroing.get():
            self.zeroing_frame.grid(row=7, column=0, padx=5, pady=5, sticky="nsew", columnspan=6)
        else:
            self.zeroing_frame.grid_forget()

    def show_differential(self):
        if self.differential.get():
            self.differential_frame.grid(row=8, column=0, padx=5, pady=5, sticky="nsew", columnspan=6)
        else:
            self.differential_frame.grid_forget()

    def show_reflexes(self):
        if self.reflexes.get():
            self.reflexes_frame.grid(row=6, column=0, padx=5, pady=5, sticky="nsew", columnspan=6)
        else:
            self.reflexes_frame.grid_forget()

    def generate(self):
        obj = self.generator_object.measurement_data
        obj['investment_zip_code'] = self.investment_zip_code_entry.get()
        obj['investment_name'] = self.investment_name_entry.get()
        obj['investment_street_and_number'] = self.investment_street_and_number_entry.get()
        obj['investment_city'] = self.investment_city_entry.get()
        obj['investors_zip_code'] = self.investors_zip_code_entry.get()
        obj['investors_name'] = self.investors_name_entry.get()
        obj['investors_street_and_number'] = self.investors_street_and_number_entry.get()
        obj['investors_city'] = self.investors_city_entry.get()
        obj['date'] = self.date.get()
        obj.update(self.zeroing_frame.building_diagram)
        obj.update(self.differential_frame.building_diagram)
        obj.update(self.reflexes_frame.reflexes_data)

        option = self.generator_object.create_directory()
        if option:
            call_dict = {
                self.front_page.get(): self.generator_object.print_front_page,
                self.differential.get(): self.generator_object.differential,
                self.zeroing.get(): self.generator_object.zeroing,
                self.insulation.get(): self.generator_object.insulation,
                self.reflexes.get(): self.generator_object.reflexes
            }

            for key in call_dict.keys():
                if key:
                    call_dict[key]()

        CTkMessagebox(message='Generowanie plików zostało zakończone', icon='check', option_1='OK')
