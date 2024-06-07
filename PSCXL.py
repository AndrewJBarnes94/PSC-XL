import tkinter as tk
from tkinter import ttk, filedialog
import openpyxl
from collections import defaultdict
import sqlite3

class PSCXL:
    def __init__(self, root):
        self.root = root
        self.root.title("PSCXL")

        # Add a button to open the file dialog
        self.open_button = tk.Button(root, text="Open Excel File", command=self.open_file)
        self.open_button.pack(pady=10)

        # Create a dropdown menu for workbook names (initially empty)
        self.workbook_selector = ttk.Combobox(root, values=[])
        self.workbook_selector.bind("<<ComboboxSelected>>", self.update_sheet_selector)
        self.workbook_selector.pack(pady=10)

        # Create a dropdown menu for sheet names (initially empty)
        self.sheet_selector = ttk.Combobox(root, values=[])
        self.sheet_selector.bind("<<ComboboxSelected>>", self.update_table)
        self.sheet_selector.pack(pady=10)

        # Add a button to read the database
        self.read_db_button = tk.Button(root, text="Read Database", command=self.read_database)
        self.read_db_button.pack(pady=10)

        # Add a Treeview widget to display the database contents
        self.db_tree = ttk.Treeview(root)

        # Define columns for the database Treeview
        self.db_tree['columns'] = ('Workbook', 'Sheet', 'Value', 'Quantity')

        # Format columns
        self.db_tree.column("#0", width=0, stretch=tk.NO)  # Hide the first column
        self.db_tree.column('Workbook', anchor=tk.W, width=200)
        self.db_tree.column('Sheet', anchor=tk.W, width=200)
        self.db_tree.column('Value', anchor=tk.W, width=250)
        self.db_tree.column('Quantity', anchor=tk.CENTER, width=80)

        # Create headings
        self.db_tree.heading("#0", text="", anchor=tk.W)
        self.db_tree.heading('Workbook', text='Workbook', anchor=tk.W)
        self.db_tree.heading('Sheet', text='Sheet', anchor=tk.W)
        self.db_tree.heading('Value', text='Value', anchor=tk.W)
        self.db_tree.heading('Quantity', text='Quantity', anchor=tk.CENTER)

        # Pack Treeview widget
        self.db_tree.pack(pady=20)

        # Initialize SQLite database
        self.conn = sqlite3.connect('pscxl.db')
        self.cursor = self.conn.cursor()
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS workbooks (
                workbook_name TEXT, 
                sheet_name TEXT, 
                value TEXT, 
                quantity INTEGER,
                PRIMARY KEY (workbook_name, sheet_name, value)
            )
        ''')
        self.conn.commit()

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

            # Store data for all sheets in the database
            workbook_base_name = self.workbook_name.split('/')[-1].split('.')[0]
            for sheet_name in self.sheet_names:
                sheet = self.wb[sheet_name]
                sheet_data = self.read_sheet(sheet)
                duplicate_counts = self.count_duplicates(sheet_data)
                for value, count in duplicate_counts.items():
                    self.cursor.execute('''
                        INSERT OR IGNORE INTO workbooks (workbook_name, sheet_name, value, quantity) 
                        VALUES (?, ?, ?, ?)
                    ''', (workbook_base_name, sheet_name, value, count))
            self.conn.commit()
            self.read_database()

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

    def read_database(self):
        # Get unique workbook names and update the workbook selector
        self.cursor.execute('SELECT DISTINCT workbook_name FROM workbooks')
        workbooks = [row[0] for row in self.cursor.fetchall()]
        self.workbook_selector['values'] = workbooks
        if workbooks:
            self.workbook_selector.current(0)
            self.update_sheet_selector()

    def update_sheet_selector(self, event=None):
        # Get the selected workbook name
        selected_workbook = self.workbook_selector.get()

        # Get unique sheet names for the selected workbook and update the sheet selector
        self.cursor.execute('SELECT DISTINCT sheet_name FROM workbooks WHERE workbook_name = ?', (selected_workbook,))
        sheets = [row[0] for row in self.cursor.fetchall()]
        self.sheet_selector['values'] = sheets
        if sheets:
            self.sheet_selector.current(0)
            self.update_table()

    def update_table(self, event=None):
        # Clear the current table
        for item in self.db_tree.get_children():
            self.db_tree.delete(item)

        # Get the selected workbook and sheet name
        selected_workbook = self.workbook_selector.get()
        selected_sheet = self.sheet_selector.get()

        # Query the database for the selected workbook and sheet
        self.cursor.execute('SELECT * FROM workbooks WHERE workbook_name = ? AND sheet_name = ?', 
                            (selected_workbook, selected_sheet))
        rows = self.cursor.fetchall()

        # Insert data into the database Treeview widget
        for row in rows:
            self.db_tree.insert(parent='', index='end', values=row)

if __name__ == "__main__":
    root = tk.Tk()
    app = PSCXL(root)
    root.mainloop()
