import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from docx import Document
import random
from datetime import datetime, timedelta


# Initialize global DataFrames
subjects_df = pd.DataFrame()
rooms_df = pd.DataFrame()
faculty_df = pd.DataFrame()
generated_timetable = []

# Function to read Excel files and return DataFrame with normalized column names
def read_excel_file(file_path):
    try:
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip().str.lower()
        return df
    except Exception as e:
        messagebox.showerror("Error", f"Error reading file: {e}")
        return None

# Function to load and display subject data
def load_subject_data():
    file_path = filedialog.askopenfilename(title="Select Subject Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        global subjects_df
        subjects_df = read_excel_file(file_path)
        if subjects_df is not None:
            display_subjects()

# Function to display subject data in Treeview
def display_subjects():
    # Clear existing items in the subject frame
    for widget in subject_frame.winfo_children():
        widget.destroy()

    global subject_selected
    subject_selected = {}

    # Create a frame for each year and its subjects
    for year in subjects_df.columns:
        year_frame = ttk.LabelFrame(subject_frame, text=year.title())
        year_frame.pack(fill=tk.X, padx=5, pady=5)
        
        for subject in subjects_df[year].dropna().tolist():
            subject_var = tk.BooleanVar()
            checkbox = tk.Checkbutton(year_frame, text=subject, variable=subject_var)
            checkbox.pack(anchor='w')
            subject_selected[f"{year}_{subject}"] = subject_var

# Function to add a subject
def add_subject():
    global subjects_df
    year = simpledialog.askstring("Input", "Enter Year:")
    subject_name = simpledialog.askstring("Input", "Enter Subject Name:")
    
    if year and subject_name:
        if year not in subjects_df.columns:
            subjects_df[year] = pd.Series(dtype=str)
        
        empty_index = subjects_df[year].isna().idxmax() if subjects_df[year].isna().any() else len(subjects_df)
        subjects_df.at[empty_index, year] = subject_name
        
        display_subjects()
        messagebox.showinfo("Success", f"Subject '{subject_name}' added to year '{year}'.")

# Function to delete selected subjects
def delete_subject():
    deleted_subjects = []
    for key, var in subject_selected.items():
        if var.get():  # If the checkbox is selected
            parts = key.split('_')
            year = parts[0]  # First part is the year
            subject = '_'.join(parts[1:])  # The rest is the subject name
            
            # Safely replace the subject with NaN in the DataFrame
            subjects_df[year].replace(subject, pd.NA, inplace=True)
            deleted_subjects.append(subject)
    
    if deleted_subjects:
        display_subjects()  # Refresh the display after deletion
        messagebox.showinfo("Success", f"Deleted subjects: {', '.join(deleted_subjects)}")
    else:
        messagebox.showwarning("Warning", "No subject selected for deletion.")

# Function to load and display room data
def load_room_data():
    file_path = filedialog.askopenfilename(title="Select Room Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        global rooms_df
        rooms_df = read_excel_file(file_path)
        if rooms_df is not None:
            display_rooms()

# Function to display room data with checkboxes
def display_rooms():
    # Clear existing items in the room frame
    for widget in room_frame.winfo_children():
        widget.destroy()

    global room_selected
    room_selected = {}

    # Create a frame for each room type and its rooms
    for room_type in rooms_df.columns:
        room_frame_type = ttk.LabelFrame(room_frame, text=room_type.title())
        room_frame_type.pack(fill=tk.X, padx=5, pady=5)
        
        for room in rooms_df[room_type].dropna().tolist():
            room_var = tk.BooleanVar()
            checkbox = tk.Checkbutton(room_frame_type, text=room, variable=room_var)
            checkbox.pack(anchor='w')
            room_selected[f"{room_type}_{room}"] = room_var

# Function to add a room
def add_room():
    room_type = simpledialog.askstring("Input", "Enter Room Type:")
    room_name = simpledialog.askstring("Input", "Enter Room Name:")
    
    if room_type and room_name:
        if room_type not in rooms_df.columns:
            rooms_df[room_type] = pd.Series(dtype=str)
        
        empty_index = rooms_df[room_type].isna().idxmax() if rooms_df[room_type].isna().any() else len(rooms_df)
        rooms_df.at[empty_index, room_type] = room_name
        
        display_rooms()
        messagebox.showinfo("Success", f"Room '{room_name}' added to type '{room_type}'.")

# Function to delete selected rooms
def delete_room():
    deleted_rooms = []
    for key, var in room_selected.items():
        if var.get():  # If the checkbox is selected
            parts = key.split('_')
            room_type = parts[0]  # First part is the room type
            room = '_'.join(parts[1:])  # The rest is the room name
            
            # Safely replace the room with NaN in the DataFrame
            rooms_df[room_type].replace(room, pd.NA, inplace=True)
            deleted_rooms.append(room)
    
    if deleted_rooms:
        display_rooms()  # Refresh the display after deletion
        messagebox.showinfo("Success", f"Deleted rooms: {', '.join(deleted_rooms)}")
    else:
        messagebox.showwarning("Warning", "No room selected for deletion.")

# Function to load and display faculty data
def load_faculty_data():
    file_path = filedialog.askopenfilename(title="Select Faculty Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        global faculty_df
        faculty_df = read_excel_file(file_path)
        if faculty_df is not None:
            display_faculty()

# Function to display faculty data and selection checkboxes
def display_faculty():
    for widget in faculty_frame.winfo_children():
        widget.destroy()  

    global faculty_selected
    faculty_selected = {}

    for idx, row in faculty_df.iterrows():
        faculty_name = row.get('faculty name', 'N/A')
        occupation = row.get('occupation', 'N/A')
        experience = row.get('experience', 'N/A')

        # Create a checkbox for each faculty member
        faculty_selected[faculty_name] = tk.BooleanVar()
        checkbox = tk.Checkbutton(faculty_frame, text=f"{faculty_name} ({occupation}, {experience} years)", variable=faculty_selected[faculty_name])
        checkbox.pack(anchor='w')

# Function to generate timetable with editable cells on double-clickdef generate_available_dates(start_date, num_days):
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from datetime import datetime, timedelta
import random

# Generate available dates
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from datetime import datetime, timedelta
import random

# Generate available dates
def generate_available_dates(start_date, num_days):
    return [(start_date + timedelta(days=i)).strftime("%Y-%m-%d") for i in range(num_days)]

# Function to get custom time slots from the user
def get_custom_time_slots(num_slots):
    time_slots = []
    for i in range(num_slots):
        time_slot = simpledialog.askstring("Time Slot", f"Enter time slot {i + 1} (e.g., 12:00 PM - 1:00 PM):")
        time_slots.append(time_slot)
    return time_slots

# Function to take exam duration, number of slots, and assign times
def get_exam_duration_and_slots():
    try:
        # Get exam duration (number of days)
        num_days = int(simpledialog.askstring("Exam Duration", "Enter exam duration (1, 2, or 3 days):"))

        if num_days not in [1, 2, 3]:
            raise ValueError("Invalid number of days.")

        # Get dates for those days
        exam_dates = []
        for i in range(num_days):
            exam_date = simpledialog.askstring("Exam Date", f"Enter date for day {i + 1} (YYYY-MM-DD):")
            exam_dates.append(exam_date)

        # Get number of slots per day (1, 2, or 3)
        num_slots = int(simpledialog.askstring("Number of Slots", "Enter number of slots per day (1, 2, or 3):"))

        if num_slots not in [1, 2, 3]:
            raise ValueError("Invalid number of slots.")

        # Get custom time slots from the user
        time_slots = get_custom_time_slots(num_slots)

        return exam_dates, time_slots
    except ValueError as ve:
        messagebox.showerror("Input Error", str(ve))
        return None, None

# Function to generate timetable with editable cells on double-click
# Function to generate timetable with unique subjects for each year
# Function to generate timetable with unique subjects for each year
# Function to generate timetable with unique subjects for each year and multiple slots per day
# Function to generate timetable with unique subjects for each year and multiple slots per day
def generate_timetable():
    # Check if dataframes are available
    if subjects_df.empty or rooms_df.empty or faculty_df.empty:
        messagebox.showwarning("Warning", "Please upload subject, room, and faculty files first.")
        return

    # Clear previous timetable if it exists
    for widget in content_frame.winfo_children():
        if isinstance(widget, ttk.Treeview):
            widget.destroy()

    global generated_timetable
    generated_timetable = []  # Reset the generated timetable

    # Get exam duration and custom time slots from user
    exam_dates, time_slots = get_exam_duration_and_slots()
    if not exam_dates or not time_slots:
        return

    columns = ("Date", "Time", "Year", "Subject", "Room Number", "Building", "Faculty")
    timetable_tree = ttk.Treeview(content_frame, columns=columns, show='headings')

    for col in columns:
        timetable_tree.heading(col, text=col)
        timetable_tree.column(col, width=100, anchor='center')

    timetable_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    selected_faculty = [name for name, selected in faculty_selected.items() if selected.get()]
    
    if not selected_faculty:
        messagebox.showwarning("Warning", "No faculty members selected!")
        return

    faculty_filtered_df = faculty_df[faculty_df['faculty name'].isin(selected_faculty)]
    
    # Track assigned slots and subjects for each year
    assigned_slots = {}
    assigned_subjects_by_year = {}  # Dictionary to track assigned subjects by year

    # Initialize a list to ensure all subjects are assigned
    for year in subjects_df.columns:
        assigned_subjects_by_year[year] = []
        subjects_list = subjects_df[year].dropna().tolist()

    # Loop through all exam dates and assign subjects to both slots
    for exam_date in exam_dates:
        for time_slot in time_slots:
            # Assign subjects for each year for this date and slot
            for year in subjects_df.columns:
                subjects_list = subjects_df[year].dropna().tolist()

                # Find the next subject that hasn't been assigned yet for this year
                subject_to_assign = None
                for subject in subjects_list:
                    if subject not in assigned_subjects_by_year[year]:
                        subject_to_assign = subject
                        break  # Assign only one subject to the current time slot

                if subject_to_assign:
                    # Assign room and faculty for this subject
                    room_index = len(assigned_slots) % len(rooms_df)
                    room = rooms_df.iloc[room_index]
                    faculty_count = 2 if room.get('capacity', 0) > 20 else 1

                    faculty_list = []
                    for j in range(faculty_count):
                        faculty_index = (len(assigned_slots) + j) % len(faculty_filtered_df)
                        faculty = faculty_filtered_df.iloc[faculty_index]
                        faculty_name = faculty.get('faculty name', 'N/A')
                        faculty_list.append(faculty_name)

                    building = room.get('building', 'N/A')
                    room_number = room.get('room number', 'N/A')
                    faculty_names = ", ".join(faculty_list)

                    # Assign the subject to the slot
                    assigned_slots[(exam_date, time_slot)] = subject_to_assign
                    assigned_subjects_by_year[year].append(subject_to_assign)  # Track assigned subject for this year

                    # Add to timetable
                    row_data = (exam_date, time_slot, year, subject_to_assign, room_number, building, faculty_names)
                    timetable_tree.insert("", tk.END, values=row_data)
                    generated_timetable.append(row_data)

    # Add editing functionality on double-click
    timetable_tree.bind("<Double-1>", lambda event: edit_cell(timetable_tree, event))

# Helper functions for time slots and exam duration remain unchanged

# Function to allow editing timetable cells on double-click
def edit_cell(tree, event):
    region = tree.identify("region", event.x, event.y)
    if region != "cell":
        return

    column = tree.identify_column(event.x)
    row_id = tree.identify_row(event.y)
    column_number = int(column.replace('#', '')) - 1

    old_value = tree.set(row_id, column_number)

    new_value = simpledialog.askstring("Edit Cell", f"Edit the value:", initialvalue=old_value)
    
    if new_value:
        tree.set(row_id, column_number, new_value)

        # Update the global generated_timetable list
        row_idx = int(row_id.strip('I')) - 1
        row_data = list(generated_timetable[row_idx])
        row_data[column_number] = new_value
        generated_timetable[row_idx] = tuple(row_data)

# Example use of the function
# Assuming subjects_df, rooms_df, faculty_df, content_frame, and faculty_selected are already defined in your project.


# Function to export timetable to a Word document
# Function to export timetable to a Word document
def export_timetable():
    if not generated_timetable:
        messagebox.showwarning("Warning", "No timetable generated yet!")
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])

    if file_path:
        try:
            doc = Document()
            doc.add_heading("Generated Exam Timetable", level=1)

            # Add table with appropriate number of columns
            table = doc.add_table(rows=1, cols=len(generated_timetable[0]))

            # Add headers
            hdr_cells = table.rows[0].cells
            headers = ("Date", "Time", "Year", "Subject", "Room Number", "Building", "Faculty")
            for i, col_name in enumerate(headers):
                hdr_cells[i].text = col_name

            # Add all rows of the generated timetable
            for row_data in generated_timetable:
                row_cells = table.add_row().cells
                for i, value in enumerate(row_data):
                    row_cells[i].text = str(value)

            # Save the document
            doc.save(file_path)
            messagebox.showinfo("Success", f"Timetable exported successfully to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Error exporting timetable: {e}")

# Tkinter UI
root = tk.Tk()
root.title("Timetable Generator")
root.geometry("800x600")

# Add scrollbars
scrollbar_y = tk.Scrollbar(root, orient=tk.VERTICAL)
scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)

scrollbar_x = tk.Scrollbar(root, orient=tk.HORIZONTAL)
scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

# Subject Frame
subject_frame = ttk.LabelFrame(root, text="Subjects")
subject_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

# Room Frame
room_frame = ttk.LabelFrame(root, text="Rooms")
room_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

# Faculty Frame
faculty_frame = ttk.LabelFrame(root, text="Faculty")
faculty_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

# Content Frame (For displaying timetable)
content_frame = ttk.LabelFrame(root, text="Generated Timetable")
content_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

# Menu Bar
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)

# File Menu
file_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="Load Subject Data", command=load_subject_data)
file_menu.add_command(label="Load Room Data", command=load_room_data)
file_menu.add_command(label="Load Faculty Data", command=load_faculty_data)
file_menu.add_separator()
file_menu.add_command(label="Exit", command=root.quit)

# Subject Menu
subject_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Subjects", menu=subject_menu)
subject_menu.add_command(label="Add Subject", command=add_subject)
subject_menu.add_command(label="Delete Subject", command=delete_subject)

# Room Menu
room_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Rooms", menu=room_menu)
room_menu.add_command(label="Add Room", command=add_room)
room_menu.add_command(label="Delete Room", command=delete_room)

# Timetable Menu
timetable_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Timetable", menu=timetable_menu)
timetable_menu.add_command(label="Generate Timetable", command=generate_timetable)
timetable_menu.add_command(label="Export Timetable", command=export_timetable)

root.mainloop()
