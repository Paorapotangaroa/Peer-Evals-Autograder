import os
from openpyxl import load_workbook

# Specify the folder containing the Excel files
folder_path = 'C:\\Users\\BYU Rental\\Downloads\\P4Evals'


listOfFiles = []
listOfErrors = []
# Loop through all files in the specified folder
for filename in os.listdir(folder_path):
    # Check if the file is an Excel file (you can add more checks for specific extensions if needed)
    try:
        # Construct the full file path
        file_path = os.path.join(folder_path, filename)
        
        # Load the workbook
        workbook = load_workbook(file_path, data_only=True)
        
        # Get the active sheet (first sheet that activates)
        sheet = workbook.active
        
        # Read the values in cells D21:G21
        values = [sheet['D21'].value, sheet['E21'].value, sheet['F21'].value, sheet['G21'].value]
        
        # Check if any of these values are 87 or less
        if any(value is not None and value <= 87 for value in values):
            listOfFiles.append(filename)
            os.startfile(file_path)
    except:
        listOfErrors.append(filename)

print("Check the following files:")
for file in listOfFiles:
    print(f"\t{file}")

print()
print("These couldn't be read:")
for file in listOfErrors:
    print(f"\t{file}")