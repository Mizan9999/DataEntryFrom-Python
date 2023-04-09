import tkinter
from tkinter import ttk
from tkinter import messagebox
import openpyxl
import os

window = tkinter.Tk()
window.title("Data Entry Form")

frame = tkinter.Frame(window)
frame.pack()


def enter_data():
    accepted = accept_var.get()
    if accepted == "Accepted":
        # User Info
        first_name = first_name_entry.get()
        last_name = last_name_entry.get()
        if first_name and last_name:
            title = title_combobox.get()
            age = age_spinbox.get()
            nationality = nationality_combobox.get()

            # course info
            num_semesters = num_semesters_spinbox.get()
            num_courses = num_courses_spinbox.get()
            registration_status = reg_status_var.get()

            filepath = r"D:\practice\01\data.xlsx"

            if not os.path.exists(filepath):
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                heading = ["First Name", "Last Name", "Title", "Age", "Nationality", "# Courses", "# Semesters",
                           "Registration Status"]
                sheet.append(heading)
                workbook.save(filepath)
            workbook = openpyxl.load_workbook(filepath)
            sheet = workbook.active
            sheet.append([first_name, last_name, title, age, nationality, num_courses, num_semesters
                          , registration_status])
            workbook.save(filepath)
            tkinter.messagebox.showinfo(title="Successful", message="Your data has been saved!")

        else:
            tkinter.messagebox.showwarning(title="Warning", message="Please write your first name and last name..")


    else:
        tkinter.messagebox.showwarning(title="Warning", message="You need to accept our terms and condition.")


country_names = ["Bangladesh", "USA", "India", "Pakistan", "Iran"]
# Saving User Info
user_info_frame = tkinter.LabelFrame(frame, text="User Information")
user_info_frame.grid(row=0, column=0, padx=20, pady=10)

first_name_label = tkinter.Label(user_info_frame, text="First Name")
first_name_label.grid(row=0, column=0)

last_name_label = tkinter.Label(user_info_frame, text="Last Name")
last_name_label.grid(row=0, column=1)

first_name_entry = tkinter.Entry(user_info_frame)
first_name_entry.grid(row=1, column=0)
last_name_entry = tkinter.Entry(user_info_frame)
last_name_entry.grid(row=1, column=1)

title = tkinter.Label(user_info_frame, text="Title")
title_combobox = ttk.Combobox(user_info_frame, values=["", "Mr.", "Ms."])
title.grid(row=0, column=2)
title_combobox.grid(row=1, column=2)

age_label = tkinter.Label(user_info_frame, text="Age")
age_spinbox = tkinter.Spinbox(user_info_frame, from_=18, to=100)
age_label.grid(row=2, column=0)
age_spinbox.grid(row=3, column=0)

nationality_label = tkinter.Label(user_info_frame, text="Nationality")
nationality_combobox = ttk.Combobox(user_info_frame, values=country_names)
nationality_label.grid(row=2, column=1)
nationality_combobox.grid(row=3, column=1)

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# saving Course Info
courses_frame = tkinter.LabelFrame(frame)
courses_frame.grid(row=1, column=0, sticky="news", pady=10, padx=20)

registered_label = tkinter.Label(courses_frame, text="Registration Status")
registered_label.grid(row=0, column=0)

reg_status_var = tkinter.StringVar(value="Not Registered")

registered_check = tkinter.Checkbutton(courses_frame, text="Current Registered", variable=reg_status_var,
                                       onvalue="Registered", offvalue="Not Registered")
registered_check.grid(row=1, column=0)

num_courses_label = tkinter.Label(courses_frame, text="# Completed Courses")
num_courses_spinbox = tkinter.Spinbox(courses_frame, from_=0, to='infinity')
num_courses_label.grid(row=0, column=1)
num_courses_spinbox.grid(row=1, column=1)

num_semesters_label = tkinter.Label(courses_frame, text="# Completed Semesters")
num_semesters_spinbox = tkinter.Spinbox(courses_frame, from_=0, to='infinity')
num_semesters_label.grid(row=0, column=2)
num_semesters_spinbox.grid(row=1, column=2)

for widget in courses_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Accept terms
terms_frame = tkinter.LabelFrame(frame, text="Terms & Conditions")
terms_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)

accept_var = tkinter.StringVar(value="Not Accepted")
terms_check = tkinter.Checkbutton(terms_frame, text="I accept the terms and condition.", variable=accept_var,
                                  onvalue="Accepted", offvalue="Not Accepted")
terms_check.grid(row=0, column=0)

# Button
button = tkinter.Button(frame, text="Enter Data", command=enter_data)
button.grid(row=3, column=0, sticky="news", pady=10, padx=20)

window.mainloop()
