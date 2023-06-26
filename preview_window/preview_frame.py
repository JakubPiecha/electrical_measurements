import customtkinter as ctk


class PreviewFrame(ctk.CTkScrollableFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.params = {'padx': 5, 'pady': 5, 'sticky': 'nsew'}
        self.data = None

    def view_data(self, data):
        row = 0
        self.data = data
        for type_data, storeys in self.data.items():

            for storey_name, rooms in storeys.items():
                self.show_storey(row, storey_name)
                row += 1

                for room, room_data in rooms.items():

                    if room_data is not None:
                        self.show_room_data(row, storey_name, room, room_data)
                        row += 1

    def show_storey(self, row, storey_name):
        storey = ctk.CTkLabel(self, text=storey_name.upper(), font=('Helvetica', 18, 'bold'))
        storey.grid(row=row, column=0, columnspan=5, **self.params)

    def show_room_data(self, row, storey_name, room, room_data):
        room_name = ctk.CTkLabel(self, width=120, text=room.upper(), font=('Helvetica', 18, 'bold'))
        room_name.grid(row=row, column=1, **self.params)

        power_sockets = ctk.CTkLabel(self, width=120, text=room_data[0], font=('Helvetica', 18, 'bold'))
        power_sockets.grid(row=row, column=2, **self.params)

        type_of_security = ctk.CTkLabel(self, width=120, text=room_data[1], font=('Helvetica', 18, 'bold'))
        type_of_security.grid(row=row, column=3, **self.params)

        protection_type = ctk.CTkLabel(self, width=120, text=room_data[2], font=('Helvetica', 18, 'bold'))
        protection_type.grid(row=row, column=4, **self.params)

        In_ = ctk.CTkLabel(self, width=120, text=room_data[3], font=('Helvetica', 18, 'bold'))
        In_.grid(row=row, column=5, **self.params)

        delete_button_room = ctk.CTkButton(self, text='Usuń', fg_color='red', width=50,
                                           command=lambda: self.delete_room(storey_name, room, row))
        delete_button_room.grid(row=row, column=6, **self.params)

    def delete_row(self, row):
        slaves = list(self.grid_slaves(row=row))
        for obj in slaves:
            obj.grid_forget()

    def delete_room(self, storey_name, room, row):
        key = list(self.data.keys())[0]
        self.delete_row(row)
        self.data[f'{key}'][storey_name].pop(room)
