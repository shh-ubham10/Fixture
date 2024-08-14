import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from PIL import Image, ImageTk
import os
import csv
from datetime import date

def load_fixture_accessory_data():
    """Load the fixture and accessory data from the CSV file."""
    filename = "fixture_data.csv"
    filepath = os.path.join(os.path.expanduser("~"), "Desktop", filename)
    
    if not os.path.exists(filepath):
        messagebox.showerror("File Not Found", f"The file {filename} does not exist.")
        return {}

    try:
        fixture_accessory_map = {}
        with open(filepath, mode='r') as file:
            reader = csv.reader(file)
            next(reader)  # Skip the header
            for row in reader:
                fixture_number, fixture_name, acc_name, acc_num, parameter, specification, inspection_instrument = row[:7]
                if fixture_number not in fixture_accessory_map:
                    fixture_accessory_map[fixture_number] = []
                fixture_accessory_map[fixture_number].append({
                    "Accessory Number": acc_num,
                    "Accessory Name": acc_name,
                    "Parameter": parameter,
                    "Specification": specification,
                    "Inspection Instrument": inspection_instrument
                })
        
        return fixture_accessory_map
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while loading the data: {str(e)}")
        return {}

def on_fixture_select(event):
    """Update the accessory combobox when a fixture number is selected."""
    selected_fixture = fixture_combobox.get()
    accessories = fixture_accessory_map.get(selected_fixture, [])
    accessory_combobox['values'] = [f"{acc['Accessory Number']} - {acc['Accessory Name']}" for acc in accessories]
    accessory_combobox.set('')  # Clear any previously selected accessory

def on_accessory_select(event):
    """Populate the TreeView with the selected accessory's details."""
    selected_fixture = fixture_combobox.get()
    selected_accessory = accessory_combobox.get()
    
    if not selected_accessory:
        return
    
    accessory_number = selected_accessory.split(" - ")[0]
    accessory_key = (selected_fixture, accessory_number)

    # Clear and populate TreeView with saved data if available
    tree.delete(*tree.get_children())  # Clear previous data
    if accessory_key in saved_observations:
        for obs in saved_observations[accessory_key]:
            tree.insert("", tk.END, values=obs)
    else:
        accessories = fixture_accessory_map.get(selected_fixture, [])
        for accessory in accessories:
            if accessory['Accessory Number'] == accessory_number:
                tree.insert("", tk.END, values=(
                    accessory['Parameter'],
                    accessory['Specification'],
                    accessory['Inspection Instrument'],
                    '',  # Observation (to be filled by user)
                    '',  # Remark (to be filled by user)
                    ''  # Default status
                ))
                break

    # Update the status button based on the populated data
    update_status_button()

def toggle_status():
    """Toggle the status between OK and Not OK, and update the TreeView."""
    new_status = 'NG' if status_button['text'] == 'OK' else 'OK'
    status_button['text'] = new_status
    status_button.config(bg='red' if new_status == 'NG' else 'green')

    # Update the status for all rows in the TreeView
    for item in tree.get_children():
        tree.set(item, column="Status", value=new_status)

def save_data_in_treeview():
    """Save the current TreeView data for the selected accessory."""
    fixture_number = fixture_combobox.get()
    selected_accessory = accessory_combobox.get()
    
    if not selected_accessory:
        messagebox.showwarning("Save Data", "Please select an accessory to save data.")
        return

    accessory_number = selected_accessory.split(" - ")[0]
    accessory_key = (fixture_number, accessory_number)

    observations = []
    for row in tree.get_children():
        observations.append(tree.item(row)["values"])

    saved_observations[accessory_key] = observations
    messagebox.showinfo("Save Data", "Data saved in TreeView. You can switch to another accessory.")

def submit_data():
    """Submit the data and save it to a single CSV file."""
    fixture_number = fixture_combobox.get()
    accessory = accessory_combobox.get()
    machine_no = machine_no_entry.get()
    operation = operation_entry.get()
    today = date_entry.get_date()

    if not all([fixture_number, accessory, machine_no, operation]):
        messagebox.showwarning("Incomplete Data", "Please fill in all the required fields.")
        return

    accessory_number, accessory_name = accessory.split(" - ")
    accessory_key = (fixture_number, accessory_number)

    save_path = os.path.join(os.path.expanduser("~"), "Desktop", "All_Accessories_Data.csv")

    try:
        file_exists = os.path.isfile(save_path)

        with open(save_path, mode='a', newline='') as file:  # Append mode
            writer = csv.writer(file)

            # Write header only if the file does not exist
            if not file_exists:
                writer.writerow(["Date", "Machine No.", "Operation", "Fixture No.", "Accessory No.", 
                                 "Accessory Name", "Parameter", "Specification", 
                                 "Inspection Instrument", "Observation", "Remark", "Status"])
            
            for obs in saved_observations.get(accessory_key, []):
                writer.writerow([today, machine_no, operation, fixture_number, accessory_number, accessory_name] + list(obs) + [status_button['text']])
        
        messagebox.showinfo("Data Submitted", f"Data has been submitted and saved to {save_path}.")
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while saving the data: {str(e)}")

def get_status_combobox():
    """Create a combobox with 'OK' and 'NG' options."""
    status_combobox = ttk.Combobox(tree, values=["OK", "NG"], state="readonly")
    return status_combobox

def edit_selected_row(event):
    """Edit the selected row directly in the TreeView."""
    selected_item = tree.selection()
    if selected_item:
        selected_item = selected_item[0]
        column = tree.identify_column(event.x)
        column_index = int(column.replace('#', '')) - 1

        def save_edit(event):
            if column_index == len(tree["columns"]) - 1:  # Status column
                status_value = edit_combobox.get()
                tree.set(selected_item, column=column, value=status_value)
                edit_combobox.destroy()
            else:
                tree.set(selected_item, column=column, value=edit_entry.get())
                edit_entry.destroy()

            update_status_button()  # Update the button based on the edited status

        x, y, width, height = tree.bbox(selected_item, column)
        if column_index == len(tree["columns"]) - 1:  # Status column
            edit_combobox = get_status_combobox()
            edit_combobox.place(x=x, y=y, width=width, height=height)
            edit_combobox.set(tree.item(selected_item, "values")[column_index])
            edit_combobox.focus()
            edit_combobox.bind("<Return>", save_edit)
            edit_combobox.bind("<FocusOut>", lambda e: edit_combobox.destroy())
        else:
            edit_entry = tk.Entry(tree)
            edit_entry.place(x=x, y=y, width=width, height=height)
            edit_entry.insert(0, tree.item(selected_item, "values")[column_index])
            edit_entry.focus()
            edit_entry.bind("<Return>", save_edit)
            edit_entry.bind("<FocusOut>", lambda e: edit_entry.destroy())

def update_status_button():
    """Update the status button based on the statuses in the TreeView."""
    all_ok = True
    for item in tree.get_children():
        status = tree.set(item, column="Status")
        if status == "NG":
            all_ok = False
            break
    
    if all_ok:
        status_button['text'] = 'OK'
        status_button.config(bg='green')
    else:
        status_button['text'] = 'NG'
        status_button.config(bg='red')

# Initialize the main application window
root = tk.Tk()
root.title("Measurement Accessory")

# Load fixture and accessory data from CSV
fixture_accessory_map = load_fixture_accessory_data()

# Dictionary to store saved observations for each accessory
saved_observations = {}

# Date input with calendar widget
tk.Label(root, text="Date:").grid(row=0, column=0, padx=10, pady=10, sticky='e')

date_entry = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
date_entry.grid(row=0, column=1, padx=10, pady=10, sticky='w')

# Machine No. input
tk.Label(root, text="Machine No.:").grid(row=1, column=0, padx=10, pady=10, sticky='e')
machine_no_entry = tk.Entry(root)
machine_no_entry.grid(row=1, column=1, padx=10, pady=10, sticky='w')

# Operation input
tk.Label(root, text="Operation:").grid(row=1, column=2, padx=10, pady=10, sticky='e')
operation_entry = tk.Entry(root)
operation_entry.grid(row=1, column=3, padx=10, pady=10, sticky='w')

# Fixture Number selection
tk.Label(root, text="Fixture No.:").grid(row=2, column=0, padx=10, pady=10, sticky='e')
fixture_combobox = ttk.Combobox(root, values=list(fixture_accessory_map.keys()))
fixture_combobox.grid(row=2, column=1, padx=10, pady=10, sticky='w')
fixture_combobox.bind('<<ComboboxSelected>>', on_fixture_select)

# Accessory Number selection
tk.Label(root, text="Accessory No.:").grid(row=2, column=2, padx=10, pady=10, sticky='e')
accessory_combobox = ttk.Combobox(root)
accessory_combobox.grid(row=2, column=3, padx=10, pady=10, sticky='w')
accessory_combobox.bind('<<ComboboxSelected>>', on_accessory_select)

# Treeview for entering details
tree = ttk.Treeview(root, columns=("Parameter", "Specification", "Inspection Instrument", 
                                   "Observation", "Remark", "Status"), show="headings")
tree.heading("Parameter", text="Parameter")
tree.heading("Specification", text="Specification")
tree.heading("Inspection Instrument", text="Inspection Instrument")
tree.heading("Observation", text="Observation")
tree.heading("Remark", text="Remark")
tree.heading("Status", text="Status")

tree.grid(row=4, column=0, columnspan=4, padx=10, pady=10, sticky='nsew')

tree.bind("<Double-1>", edit_selected_row)  # Enable double-click editing

# Status button (OK/Not OK)
status_button = tk.Button(root, text="OK", bg="green", command=toggle_status)
status_button.grid(row=6, column=0, columnspan=4, pady=10, sticky='ew')

# Save and Submit buttons
save_button = tk.Button(root, text="Save", command=save_data_in_treeview)
save_button.grid(row=7, column=0, padx=10, pady=10, sticky='ew')

submit_button = tk.Button(root, text="Submit", command=submit_data)
submit_button.grid(row=7, column=1, padx=10, pady=10, sticky='ew')

# Make sure the status button is always up-to-date with the current data
root.bind("<FocusIn>", lambda e: update_status_button())

root.mainloop()
