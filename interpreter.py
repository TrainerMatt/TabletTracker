import numpy as np
from openpyxl import load_workbook
from datetime import datetime, timedelta

class Tablet:
    def __init__(self, name: str, box_amount: int, tablets_left: int, daily) -> None:
        self.name = name
        self.box_amount = box_amount
        self.tablets_left = tablets_left
        self.daily = daily

    def __str__(self) -> str:
        return (f"{self.name} with {self.tablets_left} tablets left. {self.daily} to be taken daily. "
                f"Boxes of {self.box_amount}")

# Load the workbook and select the active worksheet
wb = load_workbook('Data.xlsx')
ws = wb.active

# Define the rows to process (corrected to use commas)
rows_to_process = [1,2,3, 4, 5, 6, 7, 8]

# List to store instances of Tablet
tablets = []

# Read data from the worksheet
print("Reading data from 'Data.xlsx':")
for row_index, row in enumerate(ws.iter_rows(min_row=min(rows_to_process), max_row=max(rows_to_process), values_only=True), start=1):
    if row_index in rows_to_process:
        # Extract cell values
        cur_name = row[0] if row[0] is not None else ''
        cur_tablets_left = row[1] if row[1] is not None else 0
        cur_box_amount = row[2] if row[2] is not None else 0
        cur_daily = row[3] if row[3] is not None else 0

        # Create an instance of Tablet
        cur_tablet = Tablet(cur_name, cur_box_amount, cur_tablets_left, cur_daily)
        tablets.append(cur_tablet)

# Print all created instances
for tablet in tablets:
    print(tablet, "\n")



def tablet_check(date: datetime, ):
