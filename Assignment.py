import pandas as pd
from datetime import datetime, timedelta

# Function
def analyze_excel_file(file_path, consecutive_days_threshold=7):
    try:
        # Reading Excel
        df = pd.read_excel(file_path)

        # Remove spaces
        df.columns = df.columns.str.strip()

        # Sets
        seven_days = set()
        short_break_printed = set()
        long_shift_printed = set()

        
        print("Query 1: ")

        for i, row in df.iterrows():
            name = row['Employee Name']
            position_id = row['Position ID']

            if name in seven_days:
                continue

            # consecutive 7 days 
            if i > 0 and name == df.at[i - 1, 'Employee Name']:
                consecutive_days = 1
                for i in range(i - 1, -1, -1):
                    if df.at[i, 'Employee Name'] == name:
                        consecutive_days += 1
                    else:
                        break
                if consecutive_days >= consecutive_days_threshold:
                    print(f"Employee: {name.title()}, Position: {position_id}")
                    seven_days.add(name)

   
        print("Query 2: ")

        employee_breaks = {}  # Dictionary to track breaks for each employee

        for i, row in df.iterrows():
            name = row['Employee Name']
            position_id = row['Position ID']

            if name in short_break_printed:
                continue

            if name in employee_breaks:
                last_time_out = employee_breaks[name]
                time_in = row['Time']

                if isinstance(time_in, str) and isinstance(last_time_out, str):
                    time_in = datetime.strptime(time_in, '%m/%d/%Y %I:%M %p')
                    last_time_out = datetime.strptime(last_time_out, '%m/%d/%Y %I:%M %p')

                    time_diff = (time_in - last_time_out).total_seconds() / 3600
                    if 1 < time_diff < 10:
                        print(f"Employee: {name.title()}, Position: {position_id}")
                        short_break_printed.add(name)
                else:
                    time_in = None

            employee_breaks[name] = row['Time Out']


        print("Query 3: ")

        for i, row in df.iterrows():
            name = row['Employee Name']
            position_id = row['Position ID']

            if name in long_shift_printed:
                continue

            # More than 14 hours
            duration_str = row['Timecard Hours (as Time)']
            if pd.notna(duration_str):
                try:
                    hours, minutes = map(int, duration_str.split(':'))
                    duration = timedelta(hours=hours, minutes=minutes)
                except ValueError:
                    # Invalid duration format 
                    duration = None
            else:
                # Handle missing values
                duration = None

            if duration is not None and duration.total_seconds() / 3600 > 14:
                print(f"Employee: {name.title()}, Position: {position_id}")
                long_shift_printed.add(name)

    except FileNotFoundError:
        print(f"File not found: {file_path}")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":

    file_path = 'Assignment_Timecard.xlsx'

    # Analyze the file
    analyze_excel_file(file_path, consecutive_days_threshold=7)