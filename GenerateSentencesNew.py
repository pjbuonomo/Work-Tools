import pandas as pd
import random
import openpyxl

# Function to remove a sheet if it exists
def remove_sheet_if_exists(workbook, sheet_name):
    if sheet_name in workbook.sheetnames:
        workbook.remove(workbook[sheet_name])

# Function to generate a sentence based on the data row
def generate_sentence(row):
    cusip = row[0]  # Column 1
    name = row[1]   # Column 1
    size = row[2]   # Column 2
    action = str(row[3]).strip().lower()  # Column 3
    try:
        price = float(row[4])  # Column 4
    except ValueError:
        return "Invalid price format"

    other_action = 'offer' if action == 'bid' else 'bid'
    other_price = price + 2 if action == 'bid' else price - 2

    formats = [
        f"{price} {action} for {size} {name} ({cusip})",
        f"{size} {name} ({cusip}) {action}ed at {price}",
        f"{size} {name} ({cusip}) - {price} {action} / {other_price} {other_action}",
        f"{name} ({cusip}) {action} at {price}",
        f"{cusip} {name} {action}ed @ {price}",
        f"{name} ({cusip}) - {action} @ {price} / {other_action} @ {other_price}",
        f"{name} ({cusip}) - {action}ed at {price} / {other_action}ed at {other_price}",
        f"{size} {name} ({cusip}) - {action}ed at {price} / {other_action}ed at {other_price}",
        f"{size} {name} ({cusip}) {action}ed @ {price}",
        f"{size} {name} ({cusip}) - {action} @ {price} / {other_action} @ {other_price}"
    ]

    return random.choice(formats)

# Load the data from the Excel file without headers
df = pd.read_excel('TrainingData.xlsx', sheet_name='Sheet1', header=None)

# Apply the function to each row
df['GeneratedSentence'] = df.apply(generate_sentence, axis=1)

# File path for the Excel file
excel_file_path = 'TrainingData.xlsx'

# Open the workbook and remove the 'Sentences' sheet if it exists
workbook = openpyxl.load_workbook(excel_file_path)
remove_sheet_if_exists(workbook, 'Sentences')
workbook.save(excel_file_path)
workbook.close()

# Save the dataframe with generated sentences to a new sheet in the same Excel file
with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
    df.to_excel(writer, sheet_name='Sentences', index=False)

print("Sentences generated and saved to Excel file.")
