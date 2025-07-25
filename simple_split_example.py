import pandas as pd

def split_semicolon_columns(input_file, output_file, columns_to_split=None):
    """
    Simple function to split semicolon-separated columns into new rows
    
    Args:
        input_file (str): Path to input Excel file
        output_file (str): Path to output Excel file  
        columns_to_split (list): Column names to split (None = auto-detect)
    """
    # Read Excel file
    df = pd.read_excel(input_file)
    print("Original data:")
    print(df)
    print(f"Original shape: {df.shape}")
    
    # Auto-detect columns with semicolons if not specified
    if columns_to_split is None:
        columns_to_split = []
        for col in df.columns:
            if df[col].astype(str).str.contains(';', na=False).any():
                columns_to_split.append(col)
    
    print(f"Columns to split: {columns_to_split}")
    
    # Split each column containing semicolons
    for col in columns_to_split:
        df[col] = df[col].astype(str).str.split(';')
        df = df.explode(col)
    
    # Clean up whitespace
    for col in columns_to_split:
        df[col] = df[col].str.strip()
    
    # Reset index
    df = df.reset_index(drop=True)
    
    print("\nProcessed data:")
    print(df)
    print(f"New shape: {df.shape}")
    
    # Save to Excel
    df.to_excel(output_file, index=False)
    print(f"Saved to: {output_file}")

# Example usage
if __name__ == "__main__":
    # Create sample data
    sample_data = {
        'ID': [1, 2, 3],
        'Name': ['John', 'Jane', 'Bob'],
        'Skills': ['Python; Java; SQL', 'HTML; CSS; JavaScript', 'Excel; PowerPoint'],
        'Department': ['IT', 'Web', 'Admin']
    }
    
    df = pd.DataFrame(sample_data)
    df.to_excel('sample_input.xlsx', index=False)
    print("Created sample_input.xlsx with semicolon-separated data")
    
    # Process the file
    split_semicolon_columns('sample_input.xlsx', 'sample_output.xlsx', ['Skills'])