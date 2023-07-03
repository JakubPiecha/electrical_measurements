import customtkinter as ctk

from main_frame import MainFrame

ctk.set_appearance_mode('dark')

ctk.set_default_color_theme('green')

class App(ctk.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.title('Dane Pomiarów Elektrycznych')

        self.main_frame = MainFrame(master=self, width=1200, height=850, corner_radius=0, fg_color="transparent")
        self.main_frame.grid(row=0, column=0, sticky="nsew")
