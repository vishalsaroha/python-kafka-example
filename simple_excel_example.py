import pandas as pd

# === READING EXCEL FILES ===

# Read an Excel file (first sheet by default)
df = pd.read_excel('input_file.xlsx')

# Read a specific sheet
df = pd.read_excel('input_file.xlsx', sheet_name='Sheet2')

# Read specific columns
df = pd.read_excel('input_file.xlsx', usecols=['Name', 'Age', 'Salary'])

# Read with custom header row
df = pd.read_excel('input_file.xlsx', header=1)  # Use row 2 as header

# === WRITING EXCEL FILES ===

# Create sample data
data = {
    'Name': ['John', 'Jane', 'Bob'],
    'Age': [25, 30, 35],
    'City': ['New York', 'London', 'Tokyo']
}
df = pd.DataFrame(data)

# Write to Excel file (basic)
df.to_excel('output_file.xlsx', index=False)

# Write to specific sheet
df.to_excel('output_file.xlsx', sheet_name='People', index=False)

# Write multiple sheets to one file
with pd.ExcelWriter('multi_sheet.xlsx') as writer:
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    df.to_excel(writer, sheet_name='Sheet2', index=False)

print("Excel operations completed successfully!")