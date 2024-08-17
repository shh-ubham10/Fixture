import tkinter as tk
from tkinter import ttk, messagebox, Menu, BooleanVar
import os
import csv
import pandas as pd
import matplotlib.pyplot as plt

def load_all_data():
    """Load all the data from the CSV file."""
    filename = "All_Accessories_Data.csv"
    current_directory = os.path.dirname(__file__)

    filepath = os.path.join(current_directory, filename)
    
    if not os.path.exists(filepath):
        messagebox.showerror("File Not Found", f"The file {filename} does not exist.")
        return [], []

    try:
        with open(filepath, mode='r') as file:
            reader = csv.reader(file)
            header = next(reader)  # Read the header
            data = []
            for row in reader:
                # Handle the case where the row has more columns than the header
                if len(row) != len(header):
                    row = row[:len(header)]  # Trim the row to match the header length
                data.append(row)
        
        return header, data
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while loading the data: {str(e)}")
        return [], []

def open_pivot_chart_window():
    """Open a new window to create a pivot chart similar to Excel."""
    pivot_window = tk.Toplevel(root)
    pivot_window.title("Pivot Chart")

    # Create labels and comboboxes for pivot options
    tk.Label(pivot_window, text="Row:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
    row_combobox = ttk.Combobox(pivot_window, values=header)
    row_combobox.grid(row=0, column=1, padx=10, pady=5, sticky="w")

    tk.Label(pivot_window, text="Column:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
    col_combobox = ttk.Combobox(pivot_window, values=header)
    col_combobox.grid(row=1, column=1, padx=10, pady=5, sticky="w")

    tk.Label(pivot_window, text="Value:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
    value_combobox = ttk.Combobox(pivot_window, values=header)
    value_combobox.grid(row=2, column=1, padx=10, pady=5, sticky="w")

    # Add a combobox for selecting the aggregation function
    tk.Label(pivot_window, text="Aggregation Function:").grid(row=3, column=0, padx=10, pady=5, sticky="e")
    agg_combobox = ttk.Combobox(pivot_window, values=["count", "sum", "mean", "min", "max"])
    agg_combobox.grid(row=3, column=1, padx=10, pady=5, sticky="w")
    agg_combobox.current(0)  # Set default to 'count'

    def generate_pivot_chart():
        """Generate the pivot chart based on the selected options."""
        row = row_combobox.get()
        col = col_combobox.get()
        value = value_combobox.get()
        aggfunc = agg_combobox.get()

        if not row or not col or not value:
            messagebox.showerror("Selection Error", "Please select a Row, Column, and Value to create the pivot chart.")
            return

        df = pd.DataFrame(data, columns=header)

        pivot_table = pd.pivot_table(
            df,
            values=value,
            index=row,
            columns=col,
            aggfunc=aggfunc,
            fill_value=0
        )

        pivot_table.plot(kind='bar', stacked=True)
        plt.title(f"Pivot Chart ({row} vs {col} with {value}, {aggfunc})")
        plt.xlabel(row)
        plt.ylabel(value)
        plt.show()

    generate_button = tk.Button(pivot_window, text="Generate Pivot Chart", command=generate_pivot_chart)
    generate_button.grid(row=4, column=0, columnspan=2, pady=20)

def filter_data(fixture_number=None, machine_numbers=None, accessory_numbers=None, dates=None):
    """Filter the data based on the selected criteria."""
    filtered_data = data
    
    if fixture_number:
        filtered_data = [row for row in filtered_data if row[fixture_col] == fixture_number]
    
    if machine_numbers:
        filtered_data = [row for row in filtered_data if row[machine_col] in machine_numbers]
    
    if accessory_numbers:
        filtered_data = [row for row in filtered_data if row[accessory_col] in accessory_numbers]
    
    if dates:
        filtered_data = [row for row in filtered_data if row[date_col] in dates]
    
    return filtered_data

def update_treeview(filtered_data):
    """Update the TreeView with the filtered data."""
    tree.delete(*tree.get_children())  # Clear the TreeView

    for row in filtered_data:
        tree.insert("", tk.END, values=row)

def on_fixture_select(event):
    """Handle the fixture selection and display the related accessories."""
    selected_fixture = fixture_combobox.get()
    filtered_data = filter_data(fixture_number=selected_fixture)
    
    # Update the TreeView with filtered data
    update_treeview(filtered_data)
    
    # Update comboboxes with available options based on the selected fixture
    update_comboboxes(filtered_data)

def update_comboboxes(filtered_data):
    """Update the comboboxes for Machine No., Accessory No., and Date based on the filtered data."""
    machine_numbers = sorted(set(row[machine_col] for row in filtered_data))
    accessory_numbers = sorted(set(row[accessory_col] for row in filtered_data))
    dates = sorted(set(row[date_col] for row in filtered_data))

    machine_combobox.config(values=machine_numbers)
    accessory_combobox.config(values=accessory_numbers)
    date_combobox.config(values=dates)

def show_column_menu(col):
    """Show a menu for selecting multiple entries when clicking on a TreeView header."""
    # Only allow the menu for specific columns
    if col not in {"Date", "Machine No.", "Accessory No."}:
        return  # Do nothing if the column is not Date, Machine No., or Accessory No.

    selected_fixture = fixture_combobox.get()
    if not selected_fixture:
        messagebox.showerror("Selection Error", "Please select a Fixture No. first.")
        return

    # Filter the data to only include the selected fixture's entries
    filtered_data = filter_data(fixture_number=selected_fixture)

    # Get unique values from the filtered data for the selected column
    col_index = header.index(col)
    unique_values = sorted(set(row[col_index] for row in filtered_data))

    menu = Menu(root, tearoff=0)
    selected_values = set()

    def toggle_selection(v, var):
        """Toggle the selection of a value."""
        if var.get():
            print(f"Attempting to add {v} to {selected_values}")
            selected_values.add(v)
            print(selected_values)
        else:
            if v in selected_values:
                print(f"Attempting to remove {v} from {selected_values}")
                selected_values.remove(v)
                print(selected_values)
            else:
                print(f"Value {v} not found in {selected_values}")
        apply_filter()


    def apply_filter():
        """Apply the filter based on selected values."""
        # If no values are selected, show all the data for the selected fixture
        if not selected_values:
            filtered_data_final = filtered_data
        else:
            filtered_data_final = [row for row in filtered_data if row[col_index] in selected_values]
        
        # Update the TreeView with the filtered data
        update_treeview(filtered_data_final)

    # Create menu items for each unique value in the column
    for value in unique_values:
        var = BooleanVar()
        print(var,value)
        menu.add_checkbutton(label=value, variable=var, 
                             command=lambda v=value, var=var: toggle_selection(v, var))

    # Display the menu at the correct location
    try:
        x, y, _, _ = tree.bbox(tree.get_children()[0], col)
    except IndexError:
        x = y = 0
    menu.post(tree.winfo_rootx() + x, tree.winfo_rooty() + y + 20)


def filter_by_column(col, values):
    """Filter TreeView data based on the selected column and multiple values."""
    selected_fixture = fixture_combobox.get()
    filtered_data = filter_data(fixture_number=selected_fixture)

    col_index = header.index(col)
    filtered_data = [row for row in filtered_data if row[col_index] in values]
    update_treeview(filtered_data)

# Initialize the main application window
root = tk.Tk()
root.title("View Accessories by Fixture")

# Load all data from the CSV file
header, data = load_all_data()

if not data:
    root.destroy()  # Exit the application if no data is loaded
else:
    # Define column indices based on headers
    fixture_col = header.index("Fixture No.")
    machine_col = header.index("Machine No.")
    accessory_col = header.index("Accessory No.")
    date_col = header.index("Date")
    
    # Fixture Number selection combobox
    tk.Label(root, text="Select Fixture No.:").grid(row=0, column=0, padx=10, pady=10, sticky='e')
    
    fixture_numbers = sorted(set(row[fixture_col] for row in data))
    fixture_combobox = ttk.Combobox(root, values=fixture_numbers)
    fixture_combobox.grid(row=0, column=1, padx=10, pady=10, sticky='w')
    fixture_combobox.bind('<<ComboboxSelected>>', on_fixture_select)
    
    # Button to open the pivot chart window
    pivot_chart_button = tk.Button(root, text="Open Pivot Chart Window", command=open_pivot_chart_window)
    pivot_chart_button.grid(row=0, column=2, padx=20, pady=20)
   
    # Create TreeView widget
    tree = ttk.Treeview(root, columns=header, show="headings")

    # Define headings and columns
    for col in header:
        tree.heading(col, text=col, command=lambda c=col: show_column_menu(c))
        tree.column(col, width=120, anchor='center')

    tree.grid(row=1, column=0, columnspan=4, padx=10, pady=10, sticky='nsew')

    # Add a scrollbar
    scrollbar = ttk.Scrollbar(root, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.grid(row=1, column=4, sticky='ns')

    # Initially, the TreeView should be empty until a fixture is selected
    update_treeview([])

root.mainloop()
