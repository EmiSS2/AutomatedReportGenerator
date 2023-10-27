import pandas as pd
from datetime import datetime, timedelta

def generate_dates(start_date, end_date):
    delta = end_date - start_date
    dates = [start_date + timedelta(days=i) for i in range(delta.days + 1)]
    return dates

def update_excel(selected_name, selected_week, statuses, template_path='Template.xlsx', output_path='Updated_Report.xlsx'):
    df = pd.read_excel(template_path, skiprows=2)  # Assuming the data starts at row 3
    print("Columns in DataFrame:", df.columns)
    
    # Find the row corresponding to the selected name
    name_idx = df[df['DATE: '] == selected_name].index[0]
    
    # Generate dates for the selected week
    start_date_str, end_date_str = selected_week.split(' - ')
    start_date = datetime.strptime(start_date_str, '%d/%m').replace(year=2023).date()
    end_date = datetime.strptime(end_date_str, '%d/%m').replace(year=2023).date()
    dates = generate_dates(start_date, end_date)
    
    # Update the row based on the selected week and statuses
    for i, date in enumerate(dates):
        date_str = date.strftime('%d/%m/%Y')
        col_name = f"Column{i+1}"
        df.loc[name_idx, col_name] = statuses[i]
    
    # Update the column names
    for i, date in enumerate(dates):
        date_str = date.strftime('%d/%m/%Y')
        col_name = f"Column{i+1}"
        df.rename(columns={col_name: date_str}, inplace=True)
        
    df.to_excel(output_path, index=False)

# Comment out the example usage if you're importing this file into another script
# selected_name = 'Ahmed'
# selected_week = '01/11 - 07/11'
# statuses = ['O', 'O', 'X', 'SICK', 'VAC', 'O', 'FRI']
# update_excel(selected_name, selected_week, statuses)
