from tkinter import Tk, Label, Button, OptionMenu, StringVar, Frame
from datetime import datetime, timedelta
from excel_handler import update_excel  # Importing the update_excel function

def generate_weeks(start_date, num_weeks=4):
    weeks = []
    for i in range(num_weeks):
        start = start_date + timedelta(days=i*7)
        end = start_date + timedelta(days=i*7 + 6)
        weeks.append(f"{start.strftime('%d/%m')} - {end.strftime('%d/%m')}")
    return weeks

def export_data():
    selected_name = name_var.get()
    selected_week = week_var.get()
    statuses = [day_vars[day].get() for day in days]  # Collecting statuses from dropdowns
    print(f"Exporting data for {selected_name} for the week {selected_week}")
    update_excel(selected_name, selected_week, statuses)  # Calling update_excel function

# Initialize Tkinter window
root = Tk()
root.title("Automated Report Generator")

# Create a dropdown for names
name_var = StringVar(root)
name_var.set("Select Name")
names = ['Ahmed', 'Ehsan', 'Haithem', 'Abdulraheem', 'Taha', 'Mawada']
name_menu = OptionMenu(root, name_var, *names)
Label(root, text="Select a Name").pack()
name_menu.pack()

# Create a dropdown for selecting the week
week_var = StringVar(root)
week_var.set("Select Week")
start_date = datetime.strptime('2023-11-01', '%Y-%m-%d').date()
weeks = generate_weeks(start_date)
week_menu = OptionMenu(root, week_var, *weeks)
Label(root, text="Select a Week").pack()
week_menu.pack()

# Create dropdowns for statuses for each day of the week
statuses = ['O', 'X', 'SICK', 'VAC', 'FRI']
days_frame = Frame(root)
days_frame.pack()

days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
day_vars = {}  # To store the StringVar for each day
for day in days:
    day_var = StringVar(days_frame)
    day_var.set("O")
    day_vars[day] = day_var  # Storing the StringVar
    Label(days_frame, text=f"{day}: ").grid(row=days.index(day), column=0)
    OptionMenu(days_frame, day_var, *statuses).grid(row=days.index(day), column=1)

# Add the export button
Button(root, text="Export", command=export_data).pack()

# Run the Tkinter event loop
root.mainloop()
