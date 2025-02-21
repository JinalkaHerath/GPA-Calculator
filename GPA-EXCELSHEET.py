import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import os


def get_grade_point(grade):
    grade_points = {
        "A+": 4.00, "A": 4.00, "A-": 3.70, "B+": 3.30, "B": 3.00, "B-": 2.70,
        "C+": 2.30, "C": 2.00, "C-": 1.70, "D+": 1.30, "D": 1.00, "E": 0.00
    }
    return grade_points.get(grade, None)


# Dictionary to store GPA per semester
semester_gpas = {}

# Function to save GPA and course data to an Excel sheet
def save_to_excel(file_name, semester, grades, credits, gpa):
    # Create a DataFrame to hold the data for the semester
    semester_data = {"Semester": [semester]}

    # Add grades and credits for each module to the dictionary
    for i, grade in enumerate(grades):
        semester_data[f"Module {i + 1} Grade"] = [grade]
    for i, credit in enumerate(credits):
        semester_data[f"Module {i + 1} Credits"] = [credit]

    # Add the Semester GPA at the end of the row
    semester_data["Semester GPA"] = [gpa]

    # Convert the data into a pandas DataFrame
    new_df = pd.DataFrame(semester_data)

    # Check if the Excel file already exists
    if os.path.exists(file_name):
        # Read the existing data from the Excel file
        existing_df = pd.read_excel(file_name, engine="openpyxl")
        # Append the new semester data
        updated_df = pd.concat([existing_df, new_df], ignore_index=True)
    else:
        # If the file doesn't exist, use the new data as the first row
        updated_df = new_df

    # Save the updated DataFrame to the Excel file
    updated_df.to_excel(file_name, index=False, engine="openpyxl")


# Function to calculate GPA and save data
def calculate_gpa():
    total_credits = 0
    total_weighted_points = 0
    semester = semester_var.get()
    file_name = file_name_var.get()  # Get the file name entered by the user

    grades = []
    credits = []

    for i in range(len(module_vars)):
        grade = grade_vars[i].get().strip().upper()
        credits_value = float(credit_vars[i].get())

        grade_point = get_grade_point(grade)
        if grade_point is None:
            messagebox.showerror("Error", f"Invalid grade entered for course {i + 1}")
            return

        grades.append(grade)
        credits.append(credits_value)

        total_credits += credits_value
        total_weighted_points += grade_point * credits_value

    if total_credits == 0:
        messagebox.showerror("Error", "Total credits cannot be zero.")
        return

    gpa = total_weighted_points / total_credits
    semester_gpas[semester] = gpa  # Store semester GPA

    # Calculate cumulative GPA
    total_gpa_weighted = sum(g * 6 for g in semester_gpas.values())  # Assuming each semester has 6 courses
    cumulative_gpa = total_gpa_weighted / (len(semester_gpas) * 6)

    # Display the calculated GPA
    gpa_label.config(text=f"Semester GPA: {round(gpa, 2)}")
    cumulative_gpa_label.config(text=f"Cumulative GPA: {round(cumulative_gpa, 2)}")

    # Save the data to Excel
    save_to_excel(file_name, semester, grades, credits, round(gpa, 2))


def clear_entries():
    semester_var.set(semester_options[0])
    file_name_var.set("GPA_Data.xlsx")  # Reset the file name to default
    for widget in course_frame.winfo_children():
        widget.destroy()
    module_vars.clear()
    grade_vars.clear()
    credit_vars.clear()
    gpa_label.config(text="Semester GPA: --")
    cumulative_gpa_label.config(text="Cumulative GPA: --")
    add_course_fields()


def add_course_fields():
    num_courses = int(num_courses_var.get())
    for i in range(num_courses):
        module_var = tk.StringVar()
        grade_var = tk.StringVar(value=grade_options[0])
        credit_var = tk.StringVar(value=credit_options[0])

        tk.Entry(course_frame, textvariable=module_var, width=10).grid(row=i, column=0, pady=5)
        ttk.Combobox(course_frame, textvariable=grade_var, values=grade_options, state="readonly", width=5).grid(row=i,
                                                                                                                 column=1,
                                                                                                                 pady=5)
        ttk.Combobox(course_frame, textvariable=credit_var, values=credit_options, state="readonly", width=5).grid(
            row=i, column=2, pady=5)

        module_vars.append(module_var)
        grade_vars.append(grade_var)
        credit_vars.append(credit_var)


# GUI Setup
root = tk.Tk()
root.title("GPA Calculator")
root.geometry("600x600")
root.configure(bg="#2C3E50")

header = tk.Label(root, text="GPA Calculator", font=("Arial", 16, "bold"), bg="#2C3E50", fg="white")
header.pack(pady=10)

semester_options = [f"Semester {i}" for i in range(1, 9)]
semester_var = tk.StringVar(value=semester_options[0])
tk.Label(root, text="Semester", font=("Arial", 12, "bold"), bg="#2C3E50", fg="white").pack()
semester_dropdown = ttk.Combobox(root, textvariable=semester_var, values=semester_options, state="readonly", width=12)
semester_dropdown.pack(pady=5)

file_name_var = tk.StringVar(value="GPA_Data.xlsx")
tk.Label(root, text="File Name", font=("Arial", 12, "bold"), bg="#2C3E50", fg="white").pack()
tk.Entry(root, textvariable=file_name_var, width=20).pack(pady=5)

num_courses_var = tk.StringVar(value="6")
tk.Label(root, text="Number of Courses", font=("Arial", 12, "bold"), bg="#2C3E50", fg="white").pack()
tk.Entry(root, textvariable=num_courses_var, width=5).pack(pady=5)
tk.Button(root, text="Set Courses", command=clear_entries, bg="#3498DB", fg="white").pack(pady=5)

course_frame = tk.Frame(root, bg="#2C3E50")
course_frame.pack()

tk.Label(course_frame, text="", font=("Arial", 12, "bold"), bg="#2C3E50", fg="white").grid(row=0, column=0)
tk.Label(course_frame, text="Grade", font=("Arial", 12, "bold"), bg="#2C3E50", fg="white").grid(row=0, column=1)
tk.Label(course_frame, text="Credits", font=("Arial", 12, "bold"), bg="#2C3E50", fg="white").grid(row=0, column=2)

grade_options = ["A+", "A", "A-", "B+", "B", "B-", "C+", "C", "C-", "D+", "D", "E"]
credit_options = ["2", "3"]

module_vars = []
grade_vars = []
credit_vars = []

add_course_fields()

gpa_button = tk.Button(root, text="Calculate GPA", font=("Arial", 12, "bold"), bg="#1ABC9C", fg="white",
                       command=calculate_gpa)
gpa_button.pack(pady=10)

clear_button = tk.Button(root, text="Clear", font=("Arial", 12, "bold"), bg="#E74C3C", fg="white",
                         command=clear_entries)
clear_button.pack(pady=10)

gpa_label = tk.Label(root, text="Semester GPA: --", font=("Arial", 20, "bold"), bg="#2C3E50", fg="white")
gpa_label.pack(pady=10)

cumulative_gpa_label = tk.Label(root, text="Cumulative GPA: --", font=("Arial", 20, "bold"), bg="#2C3E50", fg="white")
cumulative_gpa_label.pack(pady=10)

root.mainloop()
