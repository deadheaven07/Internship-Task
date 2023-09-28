import pandas as pd
from datetime import datetime, timedelta

# Function to calculate the time difference between two timestamps
def calculate_time_difference(start_time, end_time):
    start = datetime.strptime(start_time, '%I:%M %p')
    end = datetime.strptime(end_time, '%I:%M %p')
    return (end - start).total_seconds() / 3600  # Convert to hours

# Function to process the employee records from an Excel file
def process_employee_records(excel_file):
    df = pd.read_excel(excel_file)

    # Create a dictionary to store employee work data
    employee_data = {}

    for index, row in df.iterrows():
        name = row['Employee Name']
        position = row['Position Status']
        date = row['Time Out'].date()
        start_time = row['Time IN']
        end_time = row['Time Out'].time()

        if name not in employee_data:
            employee_data[name] = {'dates': set(), 'shifts': []}

        employee_data[name]['dates'].add(date)
        employee_data[name]['shifts'].append({'date': date, 'start_time': start_time, 'end_time': end_time})

    for name, data in employee_data.items():
        consecutive_days = sorted(list(data['dates']))
        if len(consecutive_days) >= 7:
            total_hours = sum(calculate_time_difference(shift['start_time'], shift['end_time']) for shift in data['shifts'])
            if total_hours > 14:
                print(f"Name: {name}, Position: {position}")

if __name__ == "__main__":
    excel_file = input("Enter the path to the Excel file: ")
    process_employee_records(excel_file)
