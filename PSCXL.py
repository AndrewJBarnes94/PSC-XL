import tkinter as tk
from tkinter import ttk
from openpyxl import load_workbook

class PSCXL:
    def __init__(self, root):
        self.root = root
        self.root.title("PSCXL")

        # Create a Treeview widget
        self.tree = ttk.Treeview(root)

        # Define columns
        self.tree['columns'] = ('Label Size', 'Text', 'Quantity')

        # Format columns
        self.tree.column("#0", width=120, minwidth=25)
        self.tree.column('Label Size', anchor=tk.W, width=120)
        self.tree.column('Text', anchor=tk.W, width=250)
        self.tree.column('Quantity', anchor=tk.CENTER, width=80)

        # Create headings
        self.tree.heading("#0", text="ID", anchor=tk.W)
        self.tree.heading('Label Size', text='Label Size', anchor=tk.W)
        self.tree.heading('Text', text='Text', anchor=tk.W)
        self.tree.heading('Quantity', text='Quantity', anchor=tk.CENTER)

        # Pack Treeview widget
        self.tree.pack(pady=20)

        # Create a frame for the input fields and labels
        self.input_frame = tk.Frame(root)
        self.input_frame.pack(pady=10)

        # Label and dropdown for Label Size
        self.label_size_label = tk.Label(self.input_frame, text="Label Size")
        self.label_size_label.grid(row=0, column=0, padx=5, pady=5)
        self.label_size_combo = ttk.Combobox(self.input_frame, values=["Wire", "Square", "Rectangle"])
        self.label_size_combo.grid(row=0, column=1, padx=5, pady=5)
        self.label_size_combo.current(0)

        # Label and entry for Text
        self.text_label = tk.Label(self.input_frame, text="Text")
        self.text_label.grid(row=0, column=2, padx=5, pady=5)
        self.text_entry = tk.Entry(self.input_frame)
        self.text_entry.grid(row=0, column=3, padx=5, pady=5)

        # Label and entry for Quantity
        self.quantity_label = tk.Label(self.input_frame, text="Quantity")
        self.quantity_label.grid(row=0, column=4, padx=5, pady=5)
        self.quantity_entry = tk.Entry(self.input_frame)
        self.quantity_entry.grid(row=0, column=5, padx=5, pady=5)

        # Add Button to add more labels
        self.add_button = tk.Button(self.input_frame, text="Add Label", command=self.add_label)
        self.add_button.grid(row=0, column=6, padx=5, pady=5)

        # Bind right-click to treeview
        self.tree.bind("<Button-3>", self.show_context_menu)

        # Create a context menu
        self.context_menu = tk.Menu(root, tearoff=0)
        self.context_menu.add_command(label="Remove Label", command=self.remove_label)

        # Load data from Excel file
        self.load_data_from_excel("24-0005E2.xlsx")

    def load_data_from_excel(self, filename):
        workbook = load_workbook(filename)
        sheet_mappings = {
            "Wire": "PANEL",
            "Square": ".5X.5 Labels",
            "Rectangle": ".5X1 Labels"
        }
        for label_size, sheet_name in sheet_mappings.items():
            sheet = workbook[sheet_name]
            for row in sheet.iter_rows(min_row=2, values_only=True):
                self.add_treeview_item(label_size, row)

    def add_treeview_item(self, label_size, row):
        next_id = len(self.tree.get_children())
        text, quantity = row
        self.tree.insert(parent='', index='end', iid=next_id, text=str(next_id + 1), values=(label_size, text, quantity))

    def add_label(self):
        label_size = self.label_size_combo.get()
        text = self.text_entry.get()
        quantity = self.quantity_entry.get()

        if label_size and text and quantity:
            # Get the next available ID
            next_id = len(self.tree.get_children())
            self.tree.insert(parent='', index='end', iid=next_id, text=str(next_id + 1), values=(label_size, text, quantity))
            # Clear entry fields
            self.text_entry.delete(0, tk.END)
            self.quantity_entry.delete(0, tk.END)

    def show_context_menu(self, event):
        # Get the row under the mouse cursor
        row_id = self.tree.identify_row(event.y)
        if row_id:
            # Select the row under the mouse cursor
            self.tree.selection_set(row_id)
            # Show the context menu
            self.context_menu.post(event.x_root, event.y_root)

    def remove_label(self):
        selected_item = self.tree.selection()[0]
        self.tree.delete(selected_item)

if __name__ == "__main__":
    root = tk.Tk()
    app = PSCXL(root)
    root.mainloop()
