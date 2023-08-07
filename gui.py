from tkinter import ttk
import tkinter as tk
from barcode_generator import barcode_generator


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title('Barcode Generator')
        self.create_widgets()

    def create_widgets(self):
        self.columnconfigure(0, weight=2)
        self.columnconfigure(1, weight=1)
        self.columnconfigure(2, weight=1)

        ttk.Label(text='How many: ').grid(row=0, column=0, padx=5, pady=5)
        ttk.Label(text='Serial Number: ').grid(row=2, column=0, padx=5, pady=5)

        total_vcmd = (self.register(self.total_entry_validate), '%P')
        total_ivcmd = (self.register(self.total_entry_on_invalid),)

        serial_vcmd = (self.register(self.serial_entry_validate), '%P')
        serial_ivcmd = (self.register(self.serial_entry_on_invalid),)

        self.total_entry = ttk.Entry(self, width=50)
        self.total_entry.config(validate='focus',
                                validatecommand=total_vcmd, invalidcommand=total_ivcmd)
        self.total_entry.grid(row=0, column=1)

        self.serial_entry = ttk.Entry(self, width=50)
        self.serial_entry.config(
            validate='focus', validatecommand=serial_vcmd, invalidcommand=serial_ivcmd)
        self.serial_entry.grid(row=2, column=1)

        self.total_error = ttk.Label(self, foreground='red')
        self.total_error.grid(row=1, column=1, sticky=tk.W)

        self.serial_error = ttk.Label(self, foreground='red')
        self.serial_error.grid(row=3, column=1, sticky=tk.W)

        self.submit_btn = ttk.Button(
            text='Save to Excel', command=lambda: self.create_barcode())
        self.submit_btn.grid(row=4, column=0, padx=5)

        self.exit_btn = ttk.Button(text='Exit', command=lambda: self.close()).grid(
            row=4, column=3, padx=5)

        self.success_msg = ttk.Label(text='')

    def close(self):
        self.destroy()

    def create_barcode(self):
        barcode_generator.create_barcode_image(
            self.total_entry.get(), self.serial_entry.get())
        self.total_entry.delete(0, tk.END)
        self.serial_entry.delete(0, tk.END)
        self.success_msg['text'] = "Done. Please check excel folder for the excel file."
        self.success_msg.grid(row=4, column=1, sticky=tk.W)

    def show_message(self, label_error, entry, error='', color='black'):
        label_error['text'] = error
        entry['foreground'] = color

    def total_entry_validate(self, value):
        if value == '' and not value.isdecimal():
            return False
        self.show_message(self.total_error, self.total_entry)
        return True

    def total_entry_on_invalid(self):
        self.show_message(self.total_error, self.total_entry,
                          'Should contain only numbers.', 'red')

    def serial_entry_validate(self, value):
        if value == '' and not value[-1].isdecimal():
            return False
        self.show_message(self.serial_error, self.serial_entry)
        return True

    def serial_entry_on_invalid(self):
        self.show_message(self.serial_error, self.serial_entry,
                          'Last serial number must be a digit.', 'red')


if __name__ == "__main__":
    app = App()
    app.mainloop()
