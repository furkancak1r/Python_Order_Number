import openpyxl

def process_excel(file_path):
    try:
        # Load the workbook and the worksheets
        workbook = openpyxl.load_workbook(file_path)
        yazilacak_sheet = workbook["INKOOL"]

        previous_value = None
        counter = 0

        # Starting from row 5 (A5), iterate through the cells in column A
        for row in range(5, yazilacak_sheet.max_row + 1):
            cell_value = yazilacak_sheet[f'A{row}'].value
            if cell_value != previous_value:
                # Reset counter when the value changes
                counter = 0
                previous_value = cell_value
            else:
                # Increment counter for repeated values
                counter += 1

            # Write the counter to the adjacent cell in column B
            yazilacak_sheet[f'B{row}'].value = counter

        # Save the workbook
        workbook.save(file_path)
        
        print("Processing completed successfully.")

    except Exception as e:
        print(f"An error occurred: {e}")

# Specify the file path
file_path = "C:\\Users\\furkan.cakir\\Desktop\\FurkanPRS\\Kodlar\\Finans & Muhasebe\\exceller\\Cariler\\OCPR-İlgili Kişiler.xlsx"

# Call the function with the specified file path
process_excel(file_path)
