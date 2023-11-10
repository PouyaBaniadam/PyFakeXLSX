import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime, timedelta
import random
from openpyxl import Workbook


def resource_path(relative_path):
    global base_path
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath("Assets")

    return os.path.join(base_path, relative_path)


class FakeDataGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PyFakeXLSX")
        self.root.iconbitmap(resource_path("Media\\fake.ico"))

        self.num_years = tk.StringVar()
        self.num_values = tk.StringVar()
        self.value_entries = []
        self.data_type_comboboxes = []
        self.range_entries = []
        self.list_entries = []

        self.label_years = ttk.Label(root, text="Enter number of years:")
        self.label_years.grid(row=0, column=0, padx=10, pady=10)

        self.entry_years = ttk.Entry(root, textvariable=self.num_years)
        self.entry_years.grid(row=0, column=1, padx=10, pady=10)

        self.label_values = ttk.Label(root, text="Enter number of values:")
        self.label_values.grid(row=1, column=0, padx=10, pady=10)

        self.entry_values = ttk.Entry(root, textvariable=self.num_values)
        self.entry_values.grid(row=1, column=1, padx=10, pady=10)

        self.button_next = ttk.Button(root, text="Next", command=self.show_value_names_page)
        self.button_next.grid(row=2, column=0, columnspan=2, pady=10)

    def show_value_names_page(self):
        self.label_years.grid_forget()
        self.entry_years.grid_forget()
        self.label_values.grid_forget()
        self.entry_values.grid_forget()
        self.button_next.grid_forget()

        self.label_value_names = ttk.Label(self.root, text="Enter value names and select data types:")
        self.label_value_names.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

        self.frame_value_entries = ttk.Frame(self.root)
        self.frame_value_entries.grid(row=1, column=0, columnspan=2, padx=10, pady=10)

        for i in range(int(self.num_values.get())):
            value_label = ttk.Label(self.frame_value_entries, text=f"Value {i + 1} name:")
            value_label.grid(row=i, column=0, padx=5, pady=5)

            value_entry = ttk.Entry(self.frame_value_entries)
            value_entry.grid(row=i, column=1, padx=5, pady=5)
            self.value_entries.append(value_entry)

            datatype_label = ttk.Label(self.frame_value_entries, text="Data Type:")
            datatype_label.grid(row=i, column=2, padx=5, pady=5)

            datatype_combobox = ttk.Combobox(self.frame_value_entries, values=["integer", "float", "string"])
            datatype_combobox.grid(row=i, column=3, padx=5, pady=5)
            datatype_combobox.set("integer")
            self.data_type_comboboxes.append(datatype_combobox)

            range_label = ttk.Label(self.frame_value_entries, text="Range (start-end):")
            range_label.grid(row=i, column=4, padx=5, pady=5)

            range_entry = ttk.Entry(self.frame_value_entries)
            range_entry.grid(row=i, column=5, padx=5, pady=5)
            self.range_entries.append(range_entry)

            list_label = ttk.Label(self.frame_value_entries, text="List (comma-separated):")
            list_label.grid(row=i, column=6, padx=5, pady=5)

            list_entry = ttk.Entry(self.frame_value_entries)
            list_entry.grid(row=i, column=7, padx=5, pady=5)
            self.list_entries.append(list_entry)

        self.button_generate = ttk.Button(self.root, text="Generate Fake Data", command=self.generate_fake_data)
        self.button_generate.grid(row=2, column=0, columnspan=2, pady=10)

    def generate_fake_data(self):
        self.num_years_value = int(self.num_years.get())
        self.value_names = [entry.get() for entry in self.value_entries]
        data_types = [entry.get() for entry in self.data_type_comboboxes]
        ranges = [entry.get() for entry in self.range_entries]
        lists = [entry.get() for entry in self.list_entries]

        fake_data = self._generate_fake_data(data_types, ranges, lists)

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

        if not file_path:
            return

        self._save_to_excel(fake_data, file_path)
        messagebox.showinfo("Success", f"Fake data saved to {file_path}")

    def _generate_fake_data(self, data_types, ranges, lists):
        end_date = datetime.now()
        start_date = end_date - timedelta(days=365 * self.num_years_value)

        fake_data = []
        current_date = start_date
        while current_date <= end_date:
            row_data = {'Date': current_date.strftime('%Y-%m-%d')}
            for name, data_type, range_str, list_str in zip(self.value_names, data_types, ranges, lists):
                row_data[name] = self._generate_value(data_type, range_str, list_str)
            fake_data.append(row_data)
            current_date += timedelta(days=1)

        return fake_data

    def _generate_value(self, data_type, range_str, list_str):
        if data_type.lower() == 'integer':
            start, end = map(int, range_str.split('-'))
            return random.randint(start, end)
        elif data_type.lower() == 'float':
            start, end = map(float, range_str.split('-'))
            return round(random.uniform(start, end), 2)
        elif data_type.lower() == 'string':
            values = [item.strip() for item in list_str.split(',')]
            return random.choice(values) if values else None
        else:
            return None

    def _save_to_excel(self, data, file_path):
        wb = Workbook()
        ws = wb.active

        headers = ['Date'] + self.value_names
        ws.append(headers)

        for row_data in data:
            row_values = [row_data['Date']] + [row_data[name] for name in self.value_names]
            ws.append(row_values)

        wb.save(file_path)


if __name__ == "__main__":
    root = tk.Tk()
    app = FakeDataGeneratorApp(root)
    root.mainloop()
