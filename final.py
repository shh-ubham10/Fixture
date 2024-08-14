import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import csv
from tkcalendar import DateEntry
from datetime import date

# Functionality 1: Manage Fixtures & Accessories
def run_functionality_1():
    def add_accessory_row(accessory_row_counter):
        """Add a new row for accessory inputs."""
        
        accessory_row_counter += 1
        num_label = tk.Label(root, text=f"Accessory Number {accessory_row_counter}:")
        num_entry = tk.Entry(root)
        name_label = tk.Label(root, text=f"Accessory Name {accessory_row_counter}:")
        name_entry = tk.Entry(root)
        num_label.grid(row=accessory_row_counter + 2, column=0, padx=10, pady=10, sticky='e')
        num_entry.grid(row=accessory_row_counter + 2, column=1, padx=10, pady=10)
        name_label.grid(row=accessory_row_counter + 2, column=2, padx=10, pady=10, sticky='e')
        name_entry.grid(row=accessory_row_counter + 2, column=3, padx=10, pady=10)
        accessory_entries.append((num_entry, name_entry))
        add_button.grid(row=accessory_row_counter + 3, column=0, columnspan=4, pady=10, sticky='ew')

    def save_data_to_csv(accessory_data, filename="fixture_data.csv", update_existing=False):
        """Save the accessory data to a CSV file with the specified format."""
        filepath = os.path.join(os.path.expanduser("~"), "Desktop", filename)
        try:
            file_exists = os.path.isfile(filepath)
            existing_data = []
            if update_existing and file_exists:
                with open(filepath, mode='r') as file:
                    reader = csv.reader(file)
                    existing_data = list(reader)
            with open(filepath, mode='w', newline='') as file:
                writer = csv.writer(file)
                if not file_exists:
                    writer.writerow(["Fixture Number", "Fixture Name", "Accessory Name", "Accessory Number", "Parameter", "Specification", "Inspection Instrument"])
                if update_existing:
                    for row in existing_data:
                        if row and row[0] != accessory_data[0]['Fixture Number']:
                            writer.writerow(row)
                for accessory in accessory_data:
                    fixture_number = accessory.get('Fixture Number', fixture_entry.get())
                    fixture_name = accessory.get('Fixture Name', fixture_name_entry.get())
                    accessory_name = accessory.get('Name')
                    accessory_number = accessory.get('Number')
                    for row in accessory['Details']:
                        writer.writerow([fixture_number, fixture_name, accessory_name, accessory_number] + list(row[3:]))
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
                if any(value == '' for value in values):
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
                next(reader)
                for row in reader:
                    fixture_number, fixture_name, acc_name, acc_num, parameter, specification, inspection_instrument = row
                    accessory = next((a for a in accessory_data if a['Number'] == acc_num), None)
                    if not accessory:
                        accessory = {'Fixture Number': fixture_number, 'Fixture Name': fixture_name, 'Number': acc_num, 'Name': acc_name, 'Details': []}
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
        fixture_numbers = list({a['Fixture Number'] for a in accessory_data})
        fixture_combobox = ttk.Combobox(details_window, values=fixture_numbers)
        fixture_combobox.pack(pady=10)
        def on_fixture_select(event):
            selected_fixture = fixture_combobox.get()
            filtered_accessories = [f"Accessory {a['Number']}" for a in accessory_data if a['Fixture Number'] == selected_fixture]
            accessory_combobox['values'] = filtered_accessories
            accessory_combobox.set('')
            for row in tree.get_children():
                tree.delete(row)
        fixture_combobox.bind('<<ComboboxSelected>>', on_fixture_select)
        tk.Label(details_window, text="Select Accessory:").pack(pady=10)
        accessory_combobox = ttk.Combobox(details_window)
        accessory_combobox.pack(pady=10)
        accessory_combobox.bind('<<ComboboxSelected>>', on_accessory_select)
        details_frame = tk.Frame(details_window)
        details_frame.pack(pady=10)
        tree = ttk.Treeview(details_frame, columns=("Serial Number", "Number", "Name", "Parameter", "Specification", "Inspection Instrument"), show="headings")
        tree.heading("Serial Number", text="Serial Number")
        tree.heading("Number", text="Accessory Number")
        tree.heading("Name", text="Accessory Name")
        tree.heading("Parameter", text="Parameter")
        tree.heading("Specification", text="Specification")
        tree.heading("Inspection Instrument", text="Inspection Instrument")
        tree.pack(fill=tk.BOTH, expand=True)
        tree.bind("<Double-1>", edit_selected_row)
        add_param_button = tk.Button(details_window, text="Add Parameter", command=add_parameter_row)
        add_param_button.pack(pady=10)
        save_button = tk.Button(details_window, text="Save Accessory Details", command=save_accessory_details)
        save_button.pack(pady=10)
        def final_submit():
            save_accessory_details()
            save_data_to_csv(accessory_data, update_existing=True)
            details_window.destroy()
        submit_button = tk.Button(details_window, text="Final Submit", command=final_submit)
        submit_button.pack(pady=10)

    def submit(accessory_data, accessory_combobox,tree,details_frame):
        """Collect data from entries and show the accessory details window."""
        accessory_data.clear()
        fixture_number = fixture_entry.get()
        fixture_name = fixture_name_entry.get()
        if not fixture_number or not fixture_name:
            messagebox.showwarning("Missing Fixture Data", "Please enter both the fixture number and name before submitting.")
            return
        for num_entry, name_entry in accessory_entries:
            num = num_entry.get()
            name = name_entry.get()
            accessory_data.append({
                'Fixture Number': fixture_number,
                'Fixture Name': fixture_name,
                'Number': num,
                'Name': name,
                'Details': [(1, num, name, '', '', '')]
            })
        open_accessory_details_window()

    root = tk.Tk()
    root.title("Fixture Frame")
    accessory_row_counter = 1
    accessory_entries = []
    accessory_data = []
    fixture_frame = tk.Frame(root, padx=10, pady=10)
    fixture_frame.grid(row=0, column=0, columnspan=4, padx=10, pady=10)
    tk.Label(fixture_frame, text="Fixture Number:").grid(row=0, column=0, padx=10, pady=10, sticky='e')
    fixture_entry = tk.Entry(fixture_frame)
    fixture_entry.grid(row=0, column=1, padx=10, pady=10)
    tk.Label(fixture_frame, text="Fixture Name:").grid(row=0, column=2, padx=10, pady=10, sticky='e')
    fixture_name_entry = tk.Entry(fixture_frame)
    fixture_name_entry.grid(row=0, column=3, padx=10, pady=10)
    num_label = tk.Label(root, text="Accessory Number 1:")
    num_entry = tk.Entry(root)
    name_label = tk.Label(root, text="Accessory Name 1:")
    name_entry = tk.Entry(root)
    num_label.grid(row=1, column=0, padx=10, pady=10, sticky='e')
    num_entry.grid(row=1, column=1, padx=10, pady=10)
    name_label.grid(row=1, column=2, padx=10, pady=10, sticky='e')
    name_entry.grid(row=1, column=3, padx=10, pady=10)
    accessory_entries.append((num_entry, name_entry))
    add_button = tk.Button(root, text="Add Accessory", command=lambda:add_accessory_row(accessory_row_counter))
    add_button.grid(row=2, column=0, columnspan=4, pady=10, sticky='ew')
    bottom_frame = tk.Frame(root)
    bottom_frame.grid(row=99, column=0, columnspan=4, sticky='s', pady=10)
    submit_button = tk.Button(bottom_frame, text="Submit", command=lambda:submit(accessory_data,accessory_combobox=None,tree=None,details_frame=None))
    submit_button.pack()
    load_button = tk.Button(bottom_frame, text="Load Previous Data", command=load_data_from_csv)
    load_button.pack(pady=10)
    root.grid_rowconfigure(0, weight=0)
    root.grid_rowconfigure(99, weight=0)
    root.grid_rowconfigure(1, weight=1)
    root.mainloop()

# Functionality 2: Measure & Inspect Accessories
def run_functionality_2():
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
                next(reader)
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
        accessory_combobox.set('')

    def on_accessory_select(event):
        """Populate the TreeView with the selected accessory's details."""
        selected_fixture = fixture_combobox.get()
        selected_accessory = accessory_combobox.get()
        if not selected_accessory:
            return
        accessory_number = selected_accessory.split(" - ")[0]
        accessory_key = (selected_fixture, accessory_number)
        tree.delete(*tree.get_children())
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
                        '',
                        '',
                        ''
                    ))
                    break
        update_status_button()

    def toggle_status():
        """Toggle the status between OK and Not OK, and update the TreeView."""
        new_status = 'NG' if status_button['text'] == 'OK' else 'OK'
        status_button['text'] = new_status
        status_button.config(bg='red' if new_status == 'NG' else 'green')
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
            with open(save_path, mode='a', newline='') as file:
                writer = csv.writer(file)
                if not file_exists:
                    writer.writerow(["Date", "Machine No.", "Operation", "Fixture No.", "Accessory No.", "Accessory Name", "Parameter", "Specification", "Inspection Instrument", "Observation", "Remark", "Status"])
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
                if column_index == len(tree["columns"]) - 1:
                    status_value = edit_combobox.get()
                    tree.set(selected_item, column=column, value=status_value)
                    edit_combobox.destroy()
                else:
                    tree.set(selected_item, column=column, value=edit_entry.get())
                    edit_entry.destroy()
                update_status_button()
            x, y, width, height = tree.bbox(selected_item, column)
            if column_index == len(tree["columns"]) - 1:
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

    root = tk.Tk()
    root.title("Measurement Accessory")
    fixture_accessory_map = load_fixture_accessory_data()
    saved_observations = {}
    tk.Label(root, text="Date:").grid(row=0, column=0, padx=10, pady=10, sticky='e')
    date_entry = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
    date_entry.grid(row=0, column=1, padx=10, pady=10, sticky='w')
    tk.Label(root, text="Machine No.:").grid(row=1, column=0, padx=10, pady=10, sticky='e')
    machine_no_entry = tk.Entry(root)
    machine_no_entry.grid(row=1, column=1, padx=10, pady=10, sticky='w')
    tk.Label(root, text="Operation:").grid(row=1, column=2, padx=10, pady=10, sticky='e')
    operation_entry = tk.Entry(root)
    operation_entry.grid(row=1, column=3, padx=10, pady=10, sticky='w')
    tk.Label(root, text="Fixture No.:").grid(row=2, column=0, padx=10, pady=10, sticky='e')
    fixture_combobox = ttk.Combobox(root, values=list(fixture_accessory_map.keys()))
    fixture_combobox.grid(row=2, column=1, padx=10, pady=10, sticky='w')
    fixture_combobox.bind('<<ComboboxSelected>>', on_fixture_select)
    tk.Label(root, text="Accessory No.:").grid(row=2, column=2, padx=10, pady=10, sticky='e')
    accessory_combobox = ttk.Combobox(root)
    accessory_combobox.grid(row=2, column=3, padx=10, pady=10, sticky='w')
    accessory_combobox.bind('<<ComboboxSelected>>', on_accessory_select)
    tree = ttk.Treeview(root, columns=("Parameter", "Specification", "Inspection Instrument", "Observation", "Remark", "Status"), show="headings")
    tree.heading("Parameter", text="Parameter")
    tree.heading("Specification", text="Specification")
    tree.heading("Inspection Instrument", text="Inspection Instrument")
    tree.heading("Observation", text="Observation")
    tree.heading("Remark", text="Remark")
    tree.heading("Status", text="Status")
    tree.grid(row=4, column=0, columnspan=4, padx=10, pady=10, sticky='nsew')
    tree.bind("<Double-1>", edit_selected_row)
    status_button = tk.Button(root, text="OK", bg="green", command=toggle_status)
    status_button.grid(row=6, column=0, columnspan=4, pady=10, sticky='ew')
    save_button = tk.Button(root, text="Save", command=save_data_in_treeview)
    save_button.grid(row=7, column=0, padx=10, pady=10, sticky='ew')
    submit_button = tk.Button(root, text="Submit", command=submit_data)
    submit_button.grid(row=7, column=1, padx=10, pady=10, sticky='ew')
    root.bind("<FocusIn>", lambda e: update_status_button())
    root.mainloop()

# Functionality 3: View Accessories Data
def run_functionality_3():
    def load_all_data():
        """Load all the data from the CSV file."""
        filename = "All_Accessories_Data.csv"
        filepath = os.path.join(os.path.expanduser("~"), "Desktop", filename)
        if not os.path.exists(filepath):
            messagebox.showerror("File Not Found", f"The file {filename} does not exist.")
            return [], []
        try:
            with open(filepath, mode='r') as file:
                reader = csv.reader(file)
                header = next(reader)
                data = [row for row in reader]
            return header, data
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while loading the data: {str(e)}")
            return [], []

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
        tree.delete(*tree.get_children())
        for row in filtered_data:
            tree.insert("", tk.END, values=row)

    def on_fixture_select(event):
        """Handle the fixture selection and display the related accessories."""
        selected_fixture = fixture_combobox.get()
        filtered_data = filter_data(fixture_number=selected_fixture)
        update_treeview(filtered_data)
        update_comboboxes(filtered_data)

    def update_comboboxes(filtered_data):
        """Update the comboboxes for Machine No., Accessory No., and Date based on the filtered data."""
        machine_numbers = sorted(set(row[machine_col] for row in filtered_data))
        accessory_numbers = sorted(set(row[accessory_col] for row in filtered_data))
        dates = sorted(set(row[date_col] for row in filtered_data))
        machine_combobox.config(values=machine_numbers)
        accessory_combobox.config(values=accessory_numbers)
        date_combobox.config(values=dates)

    def show_menu(event, column_name, filtered_data):
        """Show a dropdown menu with checkbuttons for multi-selection."""
        menu = tk.Menu(root, tearoff=0)
        if column_name == "Machine No.":
            options = sorted(set(row[machine_col] for row in filtered_data))
            current_selection = selected_machines
        elif column_name == "Accessory No.":
            options = sorted(set(row[accessory_col] for row in filtered_data))
            current_selection = selected_accessories
        elif column_name == "Date":
            options = sorted(set(row[date_col] for row in filtered_data))
            current_selection = selected_dates
        else:
            return
        for option in options:
            var = tk.BooleanVar(value=option in current_selection)
            def toggle_option(opt=option, var=var):
                if var.get():
                    current_selection.add(opt)
                else:
                    current_selection.discard(opt)
                filter_treeview()
            menu.add_checkbutton(label=option, variable=var, onvalue=True, offvalue=False, command=toggle_option)
        menu.tk_popup(event.x_root, event.y_root)

    def filter_treeview():
        """Filter the TreeView based on the current selections."""
        selected_fixture = fixture_combobox.get()
        filtered_data = filter_data(
            fixture_number=selected_fixture,
            machine_numbers=selected_machines if selected_machines else None,
            accessory_numbers=selected_accessories if selected_accessories else None,
            dates=selected_dates if selected_dates else None
        )
        update_treeview(filtered_data)

    def on_header_click(event):
        """Detect header click and show the appropriate menu."""
        region = tree.identify_region(event.x, event.y)
        if region == "heading":
            column = tree.identify_column(event.x)
            column_index = int(column[1:]) - 1
            column_name = header[column_index]
            selected_fixture = fixture_combobox.get()
            filtered_data = filter_data(fixture_number=selected_fixture)
            show_menu(event, column_name, filtered_data)

    root = tk.Tk()
    root.title("View Accessories by Fixture")
    header, data = load_all_data()
    if not data:
        root.destroy()
    else:
        fixture_col = header.index("Fixture No.")
        machine_col = header.index("Machine No.")
        accessory_col = header.index("Accessory No.")
        date_col = header.index("Date")
        selected_machines = set()
        selected_accessories = set()
        selected_dates = set()
        tk.Label(root, text="Select Fixture No.:").grid(row=0, column=0, padx=10, pady=10, sticky='e')
        fixture_numbers = sorted(set(row[fixture_col] for row in data))
        fixture_combobox = ttk.Combobox(root, values=fixture_numbers)
        fixture_combobox.grid(row=0, column=1, padx=10, pady=10, sticky='w')
        fixture_combobox.bind('<<ComboboxSelected>>', on_fixture_select)
        tree = ttk.Treeview(root, columns=header, show="headings")
        for col in header:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor='center')
        tree.grid(row=1, column=0, columnspan=4, padx=10, pady=10, sticky='nsew')
        scrollbar = ttk.Scrollbar(root, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=1, column=4, sticky='ns')
        machine_combobox = ttk.Combobox(root)
        accessory_combobox = ttk.Combobox(root)
        date_combobox = ttk.Combobox(root)
        tree.bind("<Button-1>", on_header_click)
    root.mainloop()

# Main Menu to Choose Functionality
def main_menu():
    """Create the main menu window."""
    root = tk.Tk()
    root.title("Main Menu")

    tk.Label(root, text="Choose a Functionality to Run:", font=("Helvetica", 16)).pack(pady=20)

    btn1 = tk.Button(root, text="Functionality 1: Manage Fixtures & Accessories", command=run_functionality_1, font=("Helvetica", 14))
    btn1.pack(pady=10)

    btn2 = tk.Button(root, text="Functionality 2: Measure & Inspect Accessories", command=run_functionality_2, font=("Helvetica", 14))
    btn2.pack(pady=10)

    btn3 = tk.Button(root, text="Functionality 3: View Accessories Data", command=run_functionality_3, font=("Helvetica", 14))
    btn3.pack(pady=10)

    root.mainloop()

# Start the main menu
if __name__ == "__main__":
    main_menu()
