import customtkinter as ctk
from CTkMessagebox import CTkMessagebox
from collections import defaultdict
from preview_window import PreviewWindow


class ModelFrame(ctk.CTkScrollableFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.master = master
        self.params = {'padx': 5, 'pady': 5, 'sticky': 'nsew'}
        self.name = ''
        self.type_of_security = []
        self.protection_type = []
        self.In = []
        self.building_diagram = {self.name: defaultdict(defaultdict)}
        self.floors_data = self.building_diagram[self.name]

        self.storey_name_label = ctk.CTkLabel(self, text='Nazwa kondygnacji')
        self.storey_name_label.grid(row=1, column=0, **self.params)

        self.storey_name = ctk.CTkEntry(self, placeholder_text='np. Parter')
        self.storey_name.grid(row=1, column=1, **self.params)

        self.add_storey_button = ctk.CTkButton(self, text='Dodaj kondygnację',
                                               command=lambda: self.add_storey(self.storey_name.get()))
        self.add_storey_button.grid(row=1, column=2, **self.params)

        self.preview_button = ctk.CTkButton(self, text='Podgląd', fg_color='blue',
                                            command=lambda: self.open_preview_window())
        self.preview_button.grid(row=1, column=3, **self.params)

        self.preview_window = None

    def add_storey(self, storey_name):
        last_row = self.grid_size()[-1]

        if storey_name and storey_name not in self.floors_data:
            self.floors_data[storey_name] = defaultdict(dict)
            self.storey_name.delete(0, 'end')

            room_name_Label = ctk.CTkLabel(self, text=f'{storey_name}'.capitalize())
            room_name_Label.grid(row=last_row + 1, column=0, **self.params)

            room_name = ctk.CTkEntry(self, placeholder_text='Nazwa pomieszczenia')
            room_name.grid(row=last_row + 1, column=1, **self.params)

            power_sockets = ctk.CTkEntry(self, placeholder_text='Liczba gniazdek')
            power_sockets.grid(row=last_row + 1, column=2, **self.params)

            type_of_security_option = ctk.CTkOptionMenu(self, width=100, values=self.type_of_security)
            type_of_security_option.grid(row=last_row + 1, column=3, **self.params)

            protection_type_option = ctk.CTkOptionMenu(self, width=100, values=self.protection_type)
            protection_type_option.grid(row=last_row + 1, column=4, **self.params)

            In_option = ctk.CTkOptionMenu(self, width=100, values=self.In)
            In_option.grid(row=last_row + 1, column=5, **self.params)

            add_storey_button = ctk.CTkButton(self, text='Dodaj',
                                              command=lambda: self.add_room(storey_name, room_name, power_sockets,
                                                                            type_of_security_option,
                                                                            protection_type_option,
                                                                            In_option))
            add_storey_button.grid(row=last_row + 1, column=6, **self.params)

            delete_button_storey = ctk.CTkButton(self, text='Usuń', fg_color='red', width=50,
                                                 command=lambda: self.delete_storey(storey_name, last_row + 1))
            delete_button_storey.grid(row=last_row + 1, column=7, **self.params)
        else:
            CTkMessagebox(title="Błąd", message="Kondygnacja juz istnieje lub brak nazwy", icon="cancel")

    def delete_storey(self, storey_name, row):
        key = list(self.building_diagram.keys())[0]
        for i in range(len(self.building_diagram[key][storey_name]) + 1):
            self.delete_row(row + i)
        self.building_diagram[key].pop(storey_name)

    def delete_row(self, row):
        slaves = list(self.grid_slaves(row=row))
        for obj in slaves:
            obj.grid_forget()

    def add_room(self, storey_name, room_name, power_sockets,
                 type_of_security_option, protection_type_option, In_option):

        if power_sockets.get().isdigit():
            self.floors_data[storey_name].update({
                room_name.get(): [power_sockets.get(), type_of_security_option.get(), protection_type_option.get(),
                                  In_option.get()]
            })
            room_name.delete(0, 'end')
            power_sockets.delete(0, 'end')
        else:
            CTkMessagebox(title="Błąd", message="Wprowadź liczbę gniazdek - wartość musi być liczbą", icon="cancel")

    def open_preview_window(self):
        if self.preview_window is None or not self.preview_window.winfo_exists():
            self.preview_window = PreviewWindow(self)
            self.preview_window.per_view_frame.view_data(self.building_diagram)
            self.preview_window.grab_set()
        else:
            self.preview_window.focus()
