import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
from PIL import Image, ImageTk
from datetime import date
import os
import csv

# Functionality 1: Manage Fixtures & Accessories
def run_functionality_1():
    messagebox.showinfo("Functionality 1", "Add New Fixture")
    

# Functionality 2: Measure & Inspect Accessories
def run_functionality_2():
    messagebox.showinfo("Functionality 2", "Update PM Records")

# Functionality 3: View Accessories Data
def run_functionality_3():
    messagebox.showinfo("Functionality 3", "View Past Records")

# Function to print the data for Functionality 3
def print_functionality_3_data():
    # Here you should implement the logic to print or export data from Functionality 3
    messagebox.showinfo("Print", "Printing the data for Functionality 3...")

def main_menu():
    root = tk.Tk()
    root.title("Accessory Management System")
    
    # Setting the window size
    root.geometry("580x450")
    
    # Adding the company logo at the top corner
    company_logo = Image.open("sansera_engineering_logo.jpeg")  # Make sure you have this image in your directory
    company_logo = company_logo.resize((250, 120), Image.Resampling.LANCZOS)
    company_logo_img = ImageTk.PhotoImage(company_logo)
    logo_label = tk.Label(root, image=company_logo_img)
    logo_label.image = company_logo_img
    logo_label.place(x=320, y=10)

    # Frame to hold the functionality buttons with icons
    frame = tk.Frame(root)
    frame.pack(pady=150)  # Adjust the vertical position of the buttons

    # Functionality 1
    icon1 = Image.open("Add_New_Fixture.png")  # Ensure the icon1.png exists in the directory
    icon1 = icon1.resize((50, 50), Image.Resampling.LANCZOS)
    icon1_img = ImageTk.PhotoImage(icon1)
    btn1 = tk.Button(frame, image=icon1_img, text="Add New Fixture", compound="top", command=run_functionality_1, font=("Helvetica", 12))
    btn1.image = icon1_img
    btn1.grid(row=0, column=0, padx=30)

    # Functionality 2
    icon2 = Image.open("Update_PM_Records.png")  # Ensure the icon2.png exists in the directory
    icon2 = icon2.resize((50, 50), Image.Resampling.LANCZOS)
    icon2_img = ImageTk.PhotoImage(icon2)
    btn2 = tk.Button(frame, image=icon2_img, text="Update PM Record", compound="top", command=run_functionality_2, font=("Helvetica", 12))
    btn2.image = icon2_img
    btn2.grid(row=0, column=1, padx=30)

    # Functionality 3
    icon3 = Image.open("Veiw_Past_Record.png")  # Ensure the icon3.png exists in the directory
    icon3 = icon3.resize((50, 50), Image.Resampling.LANCZOS)
    icon3_img = ImageTk.PhotoImage(icon3)
    btn3 = tk.Button(frame, image=icon3_img, text="View Past Record", compound="top", command=run_functionality_3, font=("Helvetica", 12))
    btn3.image = icon3_img
    btn3.grid(row=0, column=2, padx=30)

    # Print button at the bottom
    print_button = tk.Button(root, text="Print", command=print_functionality_3_data, font=("Helvetica", 14))
    print_button.pack(side=tk.TOP, pady=10)

    root.mainloop()

# Start the main menu
if __name__ == "__main__":
    main_menu()
