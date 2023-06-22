import customtkinter as ctk
from CTkMessagebox import CTkMessagebox


class ReflexesFrame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.params = {'padx': 5, 'pady': 5, 'sticky': 'w'}

        self.reflexes_data = {'Reflexes':[]}

        self.grounding_label = ctk.CTkLabel(self, text='Dane instalacji odgromowej')
        self.grounding_label.grid(row=1, column=0, **self.params)

        self.grounding = ctk.CTkEntry(self, placeholder_text='Liczba uziomów')
        self.grounding.grid(row=1, column=1, **self.params)

        self.ground_moisture = ctk.CTkOptionMenu(self, values=['1.4', '1.2', '2.0'])
        self.ground_moisture.grid(row=1, column=2, **self.params)

        self.required_resistance = ctk.CTkOptionMenu(self, values=['20', '1', '5', '7', '10', '30', '40', '50'])
        self.required_resistance.grid(row=1, column=3, **self.params)

        self.add_storey_button = ctk.CTkButton(self, text='Dodaj', command=self.add_reflexes_data)
        self.add_storey_button.grid(row=1, column=4, **self.params)

    def add_reflexes_data(self):
        if self.grounding.get().isdigit():
            self.reflexes_data.update(
                {'Reflexes': [self.grounding.get(), self.ground_moisture.get(), self.required_resistance.get()]})
        else:
            CTkMessagebox(title="Błąd", message="Wprowadź liczbę uziomów - wartość musi być liczbą", icon="cancel")
