import pandas as pd
import random
import openpyxl

# Function to remove a sheet if it exists
def remove_sheet_if_exists(workbook, sheet_name):
    if sheet_name in workbook.sheetnames:
        workbook.remove(workbook[sheet_name])

# Function to generate a sentence based on the data row
def generate_sentence(row):
    cusip = row[2]  # Column 'C'
    name = row[3]   # Column 'D'
    size = row[4]   # Column 'E'
    action = str(row[5]).lower()  # Convert action to string and lowercase
    try:
        price = float(row[6])  # Convert price to float
    except ValueError:
        return "Invalid price format"  # Handle invalid price format

    # Conditional logic for action and other action
    if action == 'bid':
        other_action = 'offer'
        other_price = price + 2
    elif action == 'offer':
        other_action = 'bid'
        other_price = price - 2
    else:
        return "Invalid action"  # Handle invalid action

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














































