import tkinter as tk
from tkinter import ttk, filedialog
import openpyxl
from collections import defaultdict

class PSCXL:
    def __init__(self, root):
        self.root = root
        self.root.title("PSCXL")

        # Add a button to open the file dialog
        self.open_button = tk.Button(root, text="Open Excel File", command=self.open_file)
        self.open_button.pack(pady=10)

        # Create a dropdown menu for sheet names (initially empty)
        self.sheet_selector = ttk.Combobox(root, values=[])
        self.sheet_selector.bind("<<ComboboxSelected>>", self.update_table)
        self.sheet_selector.pack(pady=10)

        # Create a Treeview widget in table mode
        self.tree = ttk.Treeview(root)

        # Define columns
        self.tree['columns'] = ('Value', 'Quantity')

        # Format columns
        self.tree.column("#0", width=0, stretch=tk.NO)  # Hide the first column
        self.tree.column('Value', anchor=tk.W, width=250)
        self.tree.column('Quantity', anchor=tk.CENTER, width=80)

        # Create headings
        self.tree.heading("#0", text="", anchor=tk.W)
        self.tree.heading('Value', text='Value', anchor=tk.W)
        self.tree.heading('Quantity', text='Quantity', anchor=tk.CENTER)

        # Pack Treeview widget
        self.tree.pack(pady=20)

        # Bind double-click to edit quantity
        self.tree.bind('<Double-1>', self.edit_quantity)

    def open_file(self):
        # Open a file dialog to select the Excel file
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if file_path:
            self.workbook_name = file_path
            self.wb = openpyxl.load_workbook(self.workbook_name)
            self.sheet_names = self.wb.sheetnames
            self.sheet_selector['values'] = self.sheet_names
            self.sheet_selector.current(0)
            self.update_table()

    def read_sheet(self, sheet):
        data = []
        for row in sheet.iter_rows(values_only=True):  # Read all rows
            data.append(list(row))  # Convert each row to a list and append to data
        return data

    def count_duplicates(self, data):
        counter = defaultdict(int)
        for row in data:
            for value in row:
                if value is not None:  # Skip empty cells
                    counter[value] += 1
        return counter

    def update_table(self, event=None):
        # Clear the current table
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Get the selected sheet name
        selected_sheet_name = self.sheet_selector.get()
        sheet = self.wb[selected_sheet_name]

        # Read sheet data and count duplicates
        self.sheet_data = self.read_sheet(sheet)
        self.duplicate_counts = self.count_duplicates(self.sheet_data)

        # Insert data into the table
        for value, count in self.duplicate_counts.items():
            self.tree.insert(parent='', index='end', values=(value, count))

    def edit_quantity(self, event):
        # Identify the region and item that was clicked
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            row_id = self.tree.identify_row(event.y)
            column_id = self.tree.identify_column(event.x)
            if column_id == '#2':  # Make sure only Quantity column is editable
                self.edit_cell(row_id, column_id)

    def edit_cell(self, row_id, column_id):
        item = self.tree.item(row_id)
        quantity = item['values'][1]  # Quantity is the second value

        # Create an entry widget at the cell's location
        x, y, width, height = self.tree.bbox(row_id, column_id)
        entry = tk.Entry(self.tree)
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, quantity)
        entry.focus()

        # Bind events to update the cell and workbook
        entry.bind("<Return>", lambda event: self.save_edit(entry, row_id))
        entry.bind("<FocusOut>", lambda event: self.save_edit(entry, row_id))

    def save_edit(self, entry, row_id):
        new_value = entry.get()
        entry.destroy()

        try:
            new_quantity = int(new_value)
        except ValueError:
            new_quantity = 1  # Default to 1 if the input is not a valid integer

        # Update the treeview
        item = self.tree.item(row_id)
        item_values = list(item['values'])
        item_values[1] = new_quantity  # Update the Quantity column
        self.tree.item(row_id, values=item_values)

        # Update the workbook
        value = item['values'][0]
        selected_sheet_name = self.sheet_selector.get()
        sheet = self.wb[selected_sheet_name]
        
        # Find and update the cell in the sheet
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value == value:
                    quantity_cell = sheet.cell(row=cell.row, column=cell.column + 1)
                    quantity_cell.value = new_quantity
                    self.wb.save(self.workbook_name)
                    return

if __name__ == "__main__":
    root = tk.Tk()
    app = PSCXL(root)
    root.mainloop()
