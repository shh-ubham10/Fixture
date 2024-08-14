import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
import csv

def add_accessory_row():
    """Add a new row for accessory inputs."""
    global accessory_row_counter
    
    # Increment row counter
    accessory_row_counter += 1
    
    # Create and place new labels and entries for the accessory
    num_label = tk.Label(root, text=f"Accessory Number {accessory_row_counter}:")
    num_entry = tk.Entry(root)
    name_label = tk.Label(root, text=f"Accessory Name {accessory_row_counter}:")
    name_entry = tk.Entry(root)

    num_label.grid(row=accessory_row_counter + 2, column=0, padx=10, pady=10, sticky='e')
    num_entry.grid(row=accessory_row_counter + 2, column=1, padx=10, pady=10)
    
    name_label.grid(row=accessory_row_counter + 2, column=2, padx=10, pady=10, sticky='e')
    name_entry.grid(row=accessory_row_counter + 2, column=3, padx=10, pady=10)
    
    # Add entries to the list for later access
    accessory_entries.append((num_entry, name_entry))

    # Move the "Add Accessory" button down
    add_button.grid(row=accessory_row_counter + 3, column=0, columnspan=4, pady=10, sticky='ew')

def save_data_to_csv(accessory_data, filename="fixture_data.csv", update_existing=False):
    """Save the accessory data to a CSV file with the specified format."""
    filepath = os.path.join(os.path.expanduser("~"), "Desktop", filename)
    
    try:
        file_exists = os.path.isfile(filepath)
        existing_data = []

        if update_existing and file_exists:
            # Read existing data
            with open(filepath, mode='r') as file:
                reader = csv.reader(file)
                existing_data = list(reader)
        
        with open(filepath, mode='w', newline='') as file:
            writer = csv.writer(file)

            if not file_exists:
                # Write CSV headers only if the file does not already exist
                writer.writerow(["Fixture Number", "Fixture Name", "Accessory Name", "Accessory Number", 
                                 "Parameter", "Specification", "Inspection Instrument"])

            if update_existing:
                # Preserve the existing data except the ones related to the currently edited fixture
                for row in existing_data:
                    if row and row[0] != accessory_data[0]['Fixture Number']:
                        writer.writerow(row)

            # Write fixture and accessory data
            for accessory in accessory_data:
                fixture_number = accessory.get('Fixture Number', fixture_entry.get())
                fixture_name = accessory.get('Fixture Name', fixture_name_entry.get())
                accessory_name = accessory.get('Name')
                accessory_number = accessory.get('Number')
                
                for row in accessory['Details']:
                    writer.writerow([fixture_number, fixture_name, accessory_name, accessory_number] + list(row[3:]))  # Adjust row order
                    
        messagebox.showinfo("Save Data", f"Data saved to {filepath}")
    
    except PermissionError:
        messagebox.showerror("Permission Error", f"Permission denied: Unable to write to {filepath}. Please ensure the file is not open and you have write permissions.")
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")

def update_treeview(accessory):
    """Update the Treeview with accessory details."""
    for row in tree.get_children():
        tree.delete(row)
    for detail in accessory['Details']:
        tree.insert("", tk.END, values=detail)

def on_accessory_select(event):
    """Handle the event when a new accessory is selected."""
    selected_index = accessory_combobox.current()
    if selected_index >= 0:
        update_treeview(accessory_data[selected_index])

def add_parameter_row():
    """Add a new row for parameter inputs in the accessory details sheet."""
    selected_index = accessory_combobox.current()
    if selected_index >= 0:
        accessory = accessory_data[selected_index]
        new_row = (len(accessory['Details']) + 1, accessory['Number'], accessory['Name'], '', '', '')
        accessory['Details'].append(new_row)
        tree.insert("", tk.END, values=new_row)

def edit_selected_row(event):
    """Edit the selected row directly in the treeview."""
    selected_item = tree.selection()
    if selected_item:
        selected_item = selected_item[0]
        column = tree.identify_column(event.x)
        column_index = int(column.replace('#', '')) - 1
    
        def save_edit(event):
            tree.set(selected_item, column=column, value=edit_entry.get())
            edit_entry.destroy()

        x, y, width, height = tree.bbox(selected_item, column)
        edit_entry = tk.Entry(details_frame)
        edit_entry.place(x=x, y=y, width=width, height=height)
        edit_entry.insert(0, tree.item(selected_item, "values")[column_index])
        edit_entry.focus()
        edit_entry.bind("<Return>", save_edit)
        edit_entry.bind("<FocusOut>", lambda e: edit_entry.destroy())

def save_accessory_details():
    """Save the current accessory details."""
    selected_index = accessory_combobox.current()
    if selected_index >= 0:
        details = []
        for row in tree.get_children():
            values = tree.item(row)["values"]
            if any(value == '' for value in values):  # Check if any value is empty
                messagebox.showwarning("Incomplete Details", "Please fill in all the details before saving.")
                return
            details.append(values)
        accessory_data[selected_index]['Details'] = details
        accessory_data[selected_index]['Fixture Number'] = fixture_entry.get()
        accessory_data[selected_index]['Fixture Name'] = fixture_name_entry.get()
        messagebox.showinfo("Accessory Saved", f"Details for Accessory {accessory_data[selected_index]['Number']} saved successfully.")

def load_data_from_csv(filename="fixture_data.csv"):
    """Load the accessory data from a CSV file and populate the application."""
    filepath = os.path.join(os.path.expanduser("~"), "Desktop", filename)
    
    if not os.path.exists(filepath):
        messagebox.showerror("File Not Found", f"The file {filename} does not exist.")
        return

    try:
        with open(filepath, mode='r') as file:
            reader = csv.reader(file)
            next(reader)  # Skip the header

            for row in reader:
                fixture_number, fixture_name, acc_name, acc_num, parameter, specification, inspection_instrument = row
                
                # Check if accessory already exists in accessory_data
                accessory = next((a for a in accessory_data if a['Number'] == acc_num), None)
                if not accessory:
                    accessory = {
                        'Fixture Number': fixture_number,
                        'Fixture Name': fixture_name,
                        'Number': acc_num,
                        'Name': acc_name,
                        'Details': []
                    }
                    accessory_data.append(accessory)

                accessory['Details'].append((1, acc_num, acc_name, parameter, specification, inspection_instrument))

            fixture_entry.delete(0, tk.END)
            fixture_entry.insert(0, fixture_number)
            fixture_name_entry.delete(0, tk.END)
            fixture_name_entry.insert(0, fixture_name)

        open_accessory_details_window()
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while loading the data: {str(e)}")

def open_accessory_details_window():
    """Open the accessory details window with the loaded data."""
    global accessory_combobox, tree, details_frame, fixture_combobox

    details_window = tk.Toplevel(root)
    details_window.title("Accessory Details")

    tk.Label(details_window, text="Select Fixture:").pack(pady=10)

    # Create a combobox to select fixture
    fixture_numbers = list({a['Fixture Number'] for a in accessory_data})
    fixture_combobox = ttk.Combobox(details_window, values=fixture_numbers)
    fixture_combobox.pack(pady=10)

    # Function to update accessory combobox based on selected fixture
    def on_fixture_select(event):
        selected_fixture = fixture_combobox.get()
        filtered_accessories = [f"Accessory {a['Number']}" for a in accessory_data if a['Fixture Number'] == selected_fixture]
        accessory_combobox['values'] = filtered_accessories
        accessory_combobox.set('')  # Clear any previously selected accessory
        for row in tree.get_children():
            tree.delete(row)  # Clear previous details in treeview

    fixture_combobox.bind('<<ComboboxSelected>>', on_fixture_select)

    tk.Label(details_window, text="Select Accessory:").pack(pady=10)

    # Create a combobox to select accessory
    accessory_combobox = ttk.Combobox(details_window)
    accessory_combobox.pack(pady=10)
    accessory_combobox.bind('<<ComboboxSelected>>', on_accessory_select)

    # Frame for treeview and buttons
    details_frame = tk.Frame(details_window)
    details_frame.pack(pady=10)

    # Create a treeview for displaying details
    tree = ttk.Treeview(details_frame, columns=("Serial Number", "Number", "Name", "Parameter", "Specification", "Inspection Instrument"), show="headings")
    tree.heading("Serial Number", text="Serial Number")
    tree.heading("Number", text="Accessory Number")
    tree.heading("Name", text="Accessory Name")
    tree.heading("Parameter", text="Parameter")
    tree.heading("Specification", text="Specification")
    tree.heading("Inspection Instrument", text="Inspection Instrument")

    tree.pack(fill=tk.BOTH, expand=True)
    tree.bind("<Double-1>", edit_selected_row)

    # Add a button to add parameters
    add_param_button = tk.Button(details_window, text="Add Parameter", command=add_parameter_row)
    add_param_button.pack(pady=10)

    # Add a button to save the accessory details
    save_button = tk.Button(details_window, text="Save Accessory Details", command=save_accessory_details)
    save_button.pack(pady=10)

    # Add a button to submit all details
    def final_submit():
        save_accessory_details()
        save_data_to_csv(accessory_data, update_existing=True)
        details_window.destroy()

    submit_button = tk.Button(details_window, text="Final Submit", command=final_submit)
    submit_button.pack(pady=10) 

def submit():
    """Collect data from entries and show the accessory details window."""
    global accessory_data, accessory_combobox, tree, details_frame

    # Clear previous accessory data to avoid duplication
    accessory_data.clear()

    fixture_number = fixture_entry.get()
    fixture_name = fixture_name_entry.get()

    # Debug: Print the fixture number and name
    print(f"Submitting fixture number: {fixture_number}, fixture name: {fixture_name}")

    # Check if fixture number or name is empty
    if not fixture_number or not fixture_name:
        messagebox.showwarning("Missing Fixture Data", "Please enter both the fixture number and name before submitting.")
        return

    # Collect accessory data
    for num_entry, name_entry in accessory_entries:
        num = num_entry.get()
        name = name_entry.get()
        accessory_data.append({
            'Fixture Number': fixture_number,  # Add fixture number to each accessory
            'Fixture Name': fixture_name,      # Add fixture name to each accessory
            'Number': num,
            'Name': name,
            'Details': [(1, num, name, '', '', '')]  # Add an initial blank row for details
        })

    # Debug: Print the collected accessory data
    print(f"Accessory data: {accessory_data}")

    # Open accessory details window
    open_accessory_details_window()

# Initialize the main application window
root = tk.Tk()
root.title("Fixture Frame")

# Initialize row counter for accessories and accessory data list
accessory_row_counter = 1
accessory_entries = []  # List to keep track of accessory entries
accessory_data = []  # List to store accessory details

# Frame for fixture data
fixture_frame = tk.Frame(root, padx=10, pady=10)
fixture_frame.grid(row=0, column=0, columnspan=4, padx=10, pady=10)

# Fixture Number
tk.Label(fixture_frame, text="Fixture Number:").grid(row=0, column=0, padx=10, pady=10, sticky='e')
fixture_entry = tk.Entry(fixture_frame)
fixture_entry.grid(row=0, column=1, padx=10, pady=10)

# Fixture Name
tk.Label(fixture_frame, text="Fixture Name:").grid(row=0, column=2, padx=10, pady=10, sticky='e')
fixture_name_entry = tk.Entry(fixture_frame)
fixture_name_entry.grid(row=0, column=3, padx=10, pady=10)

# Initial row for the first accessory
num_label = tk.Label(root, text="Accessory Number 1:")
num_entry = tk.Entry(root)
name_label = tk.Label(root, text="Accessory Name 1:")
name_entry = tk.Entry(root)

num_label.grid(row=1, column=0, padx=10, pady=10, sticky='e')
num_entry.grid(row=1, column=1, padx=10, pady=10)

name_label.grid(row=1, column=2, padx=10, pady=10, sticky='e')
name_entry.grid(row=1, column=3, padx=10, pady=10)

# Add entries to the list
accessory_entries.append((num_entry, name_entry))

# Add Accessory Button
add_button = tk.Button(root, text="Add Accessory", command=add_accessory_row)
add_button.grid(row=2, column=0, columnspan=4, pady=10, sticky='ew')

# Fixed frame for the submit button at the bottom
bottom_frame = tk.Frame(root)
bottom_frame.grid(row=99, column=0, columnspan=4, sticky='s', pady=10)
submit_button = tk.Button(bottom_frame, text="Submit", command=submit)
submit_button.pack()

# Load CSV Data Button
load_button = tk.Button(bottom_frame, text="Load Previous Data", command=load_data_from_csv)
load_button.pack(pady=10)

# Configure grid rows
root.grid_rowconfigure(0, weight=0)  # Fixture frame row
root.grid_rowconfigure(99, weight=0) # Bottom frame row

# Allow the accessory rows and the "Add Accessory" button to expand
root.grid_rowconfigure(1, weight=1)  # Accessory rows and button area will expand as needed

# Start the Tkinter event loop
root.mainloop()
