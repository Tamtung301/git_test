import openpyxl
from datetime import datetime, timedelta

# Ask the user for input
job_price = float(input("Enter rodent trapout price: "))
job_address = input("Enter job address: ")
technician_name = input("Enter technician name: ")
start_date_str = input("Enter start date (in format YYYY-MM-DD): ")
start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
snap_trap_qty = int(input("Enter snap trap quantity: "))
snap_trap_locations = input("Enter snap trap locations (comma-separated): ")
exclusion_work = input("Will you be doing exclusion work? (y/n): ")

# Calculate end date
end_date = start_date + timedelta(days=30)
if exclusion_work.lower() == 'y':
    end_date += timedelta(days=60)
    note = "Exclusion work will be done"
else:
    note = ""

# Open Excel file and add new row
wb = openpyxl.load_workbook("pest_control_jobs.xlsx")
ws = wb.active
new_row = [job_price, job_address, technician_name, start_date, snap_trap_qty, snap_trap_locations, end_date, note]
ws.append(new_row)
wb.save("pest_control_jobs.xlsx")
