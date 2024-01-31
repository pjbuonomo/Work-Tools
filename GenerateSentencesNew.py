import pandas as pd
import random
import openpyxl

# Function to remove a sheet if it exists
def remove_sheet_if_exists(workbook, sheet_name):
    if sheet_name in workbook.sheetnames:
        workbook.remove(workbook[sheet_name])

# Function to generate a sentence based on the data row
def generate_sentence(row):
    cusip = str(row[0]).strip()  # Column 1 (CUSIP)
    name = str(row[1]).strip()   # Column 1 (Name)
    size = row[2]   # Column 2
    action = str(row[3]).strip().lower()  # Column 3
    price = row[4]  # Column 4 (Price)

    # Generate a random other_price between 98.20 and 100.10
    other_price = round(random.uniform(98.20, 100.10), 2)

    # Define sentence formats
    formats_with_size = [
        f"{price} {action} for {size} {name} ({cusip})",
        f"{size} {name} ({cusip}) offered at {price}" if action == 'bid' else f"{size} {name} ({cusip}) bid at {price}",
        f"{size} {name} ({cusip}) - {price} bid / {other_price} offer",
        f"{size} {name} ({cusip}) - bid at {price} / offered at {other_price}",
        f"{size} {name} ({cusip}) - bid @ {price} / offered @ {other_price}",
        f"{size} {name} ({cusip}) - bid @ {price} / offer @ {other_price}",
        f"{size} {name} ({cusip}) {action}ed @ {price}",
        f"{size} {name} ({cusip}) - {action} @ {price} / {other_action} @ {other_price}"
    ]

    formats_without_size = [
        f"{name} ({cusip}) bid at {price}",
        f"{cusip} {name} offered @ {price}" if action == 'offer' else f"{cusip} {name} bid @ {price}",
        f"{cusip} {name} offer @ {price}" if action == 'offer' else f"{cusip} {name} bid @ {price}",
        f"{name} ({cusip}) offered at {price}"
    ]

    # Choose the appropriate list of formats based on size
    if size == 0:
        formats = formats_without_size
    else:
        formats = formats_with_size + formats_without_size

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
