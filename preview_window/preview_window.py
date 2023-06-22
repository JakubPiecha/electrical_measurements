import customtkinter as ctk

from preview_window.preview_frame import PreviewFrame


class PreviewWindow(ctk.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.geometry("850x500")
        self.title('Podgląd danych')

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.per_view_frame = PreviewFrame(master=self, corner_radius=0, fg_color="transparent")
        self.per_view_frame.grid(row=0, column=0, sticky="nsew")
