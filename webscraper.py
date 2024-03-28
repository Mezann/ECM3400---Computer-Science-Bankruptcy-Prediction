import os
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

# Name of the company
name = input("Enter the name of the company: ")

# Stock symbol
sym = input("Enter the symbol for the company: ")

# URL of the webpage
url = "https://stockanalysis.com/stocks/" + sym + "/financials/ratios/"

# Send a GET request to the URL
response = requests.get(url)

# Parse the HTML content of the webpage
soup = BeautifulSoup(response.text, 'html.parser')

# Find all <td> tags on the page
td_tags = soup.find_all('td')

# Create a new Workbook
wb = Workbook()
ws = wb.active

# Write the "Name" column
ws.cell(row=1, column=1, value="Name")
for i in range(2,13):
    ws.cell(row=i, column=1, value=name)  # Write the company name in the second row

# Initialize counters
row_counter = 1
column_counter = 2

# Write the content of each <td> tag to the Excel sheet
for idx, td_tag in enumerate(td_tags, start=1):
    # Skip every 13th <td> tag
    if idx % 13 == 0:
        continue
    
    td_content = td_tag.get_text(strip=True)
    ws.cell(row=row_counter, column=column_counter, value=td_content)
    
    # Move to the next row
    row_counter += 1
    
    # If we've reached the 12th row, move to the next column and reset the row counter
    if row_counter > 12:
        row_counter = 1
        column_counter += 1

# Add a column for years starting from 2023 to 2013
years = range(2022, 2012, -1)
for idx, year in enumerate(years, start=3):  # Start from 2 to place 2023 on the second row
    ws.cell(row=idx, column=column_counter, value=str(year))

# Add "Year" to the first row of the column
ws.cell(row=1, column=column_counter).value = "Year"
ws.cell(row=2, column=column_counter).value = "Current"

# Create the subfolder if it doesn't exist
subfolder = "Dataset"
if not os.path.exists(subfolder):
    os.makedirs(subfolder)

# Save the Excel file into the subfolder with the company name
file_path = os.path.join(subfolder, name + "_Ratios.xlsx")
wb.save(file_path)