import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import openpyxl
import os

# Create UI
window = tk.Tk()
window.title("Data Entry Form")

# Function to save data
def save_data():
    firstname = firstname_entry.get()
    lastname = lastname_entry.get()
    age = age_spinbox.get()
    gender = gender_combobox.get()
    phone = phone_entry.get()
    email = email_entry.get()
    city = city_entry.get()
    state = state_entry.get()
    country = country_entry.get()
    registration_status = registration_status_combobox.get()
    num_courses = num_courses_spinbox.get()
    num_credits = num_credits_spinbox.get()

    # Get the directory of the current script
    current_directory = os.path.dirname(os.path.realpath(__file__))

    # Construct the file path relative to the current directory
    filepath = os.path.join(current_directory, "data.xlsx")

    if os.path.exists(filepath):
        # Open the existing Excel file
        wb = openpyxl.load_workbook(filepath)
        ws = wb.active
    else:
        # Create a new Excel file if it doesn't exist
        wb = openpyxl.Workbook()
        ws = wb.active
        # Write headers if it's a new file
        headers = ["First Name", "Last Name", "Age", "Gender", "Phone", "Email", "City", "State", "Country",
                   "Registration Status", "Number of Courses", "Number of Credits"]
        ws.append(headers)

    # Write data
    data = [firstname, lastname, age, gender, phone, email, city, state, country, registration_status,
            num_courses, num_credits]
    ws.append(data)

    # Save the file
    wb.save(filepath)

    # Display a message box indicating successful save
    messagebox.showinfo("Success", "Data saved successfully!")


# Create a frame for personal information
personal_frame = ttk.LabelFrame(window, text="Personal Information")
personal_frame.grid(row=0, column=0, padx=10, pady=10)

# Label and Entry for First Name
tk.Label(personal_frame, text="First Name:").grid(row=0, column=0)
firstname_entry = tk.Entry(personal_frame)
firstname_entry.grid(row=0, column=1)

# Label and Entry for Last Name
tk.Label(personal_frame, text="Last Name:").grid(row=1, column=0)
lastname_entry = tk.Entry(personal_frame)
lastname_entry.grid(row=1, column=1)

# Label and Spinbox for Age
tk.Label(personal_frame, text="Age:").grid(row=2, column=0)
age_spinbox = tk.Spinbox(personal_frame, from_=0, to=120)
age_spinbox.grid(row=2, column=1)

# Label and Combobox for Gender
tk.Label(personal_frame, text="Gender:").grid(row=3, column=0)
gender_combobox = ttk.Combobox(personal_frame, values=["Male", "Female", "Other"])
gender_combobox.grid(row=3, column=1)

# Label and Entry for Phone
tk.Label(personal_frame, text="Phone:").grid(row=4, column=0)
phone_entry = tk.Entry(personal_frame)
phone_entry.grid(row=4, column=1)

# Label and Entry for Email
tk.Label(personal_frame, text="Email:").grid(row=5, column=0)
email_entry = tk.Entry(personal_frame)
email_entry.grid(row=5, column=1)

# Label and Entry for City
tk.Label(personal_frame, text="City:").grid(row=6, column=0)
city_entry = tk.Entry(personal_frame)
city_entry.grid(row=6, column=1)

# Label and Entry for State
tk.Label(personal_frame, text="State:").grid(row=7, column=0)
state_entry = tk.Entry(personal_frame)
state_entry.grid(row=7, column=1)

# Label and Entry for Country
tk.Label(personal_frame, text="Country:").grid(row=8, column=0)
country_entry = tk.Entry(personal_frame)
country_entry.grid(row=8, column=1)

# Create a frame for educational information
educational_frame = ttk.LabelFrame(window, text="Educational Information")
educational_frame.grid(row=1, column=0, padx=10, pady=10)

# Label and Combobox for Registration Status
tk.Label(educational_frame, text="Registration Status:").grid(row=0, column=0)
registration_status_combobox = ttk.Combobox(educational_frame, values=["Applied", "Enrolled", "Graduated", "Withdrawn"])
registration_status_combobox.grid(row=0, column=1)

# Label and Spinbox for Number of Courses
tk.Label(educational_frame, text="Number of registered courses:").grid(row=1, column=0)
num_courses_spinbox = tk.Spinbox(educational_frame, from_=0, to=5)
num_courses_spinbox.grid(row=1, column=1)

# Label and Spinbox for Number of Credits
tk.Label(educational_frame, text="Number of credits obtained:").grid(row=2, column=0)
num_credits_spinbox = tk.Spinbox(educational_frame, from_=0, to=200)
num_credits_spinbox.grid(row=2, column=1)

# Button to save data
save_button = tk.Button(window, text="Save Data", command=save_data)
save_button.grid(row=2, column=0, columnspan=2, pady=10)

# Start Tkinter event loop
window.mainloop()
