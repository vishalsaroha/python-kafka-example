import pandas as pd
import numpy as np

def split_semicolon_columns(df, columns_to_split=None, separator=';'):
    """
    Split columns containing semicolons into new rows while keeping other column values the same
    
    Args:
        df (pandas.DataFrame): Input DataFrame
        columns_to_split (list): List of column names to split. If None, splits all columns containing separator
        separator (str): Character to split on (default: ';')
    
    Returns:
        pandas.DataFrame: DataFrame with split rows
    """
    if df.empty:
        return df
    
    # If no columns specified, find all columns that contain the separator
    if columns_to_split is None:
        columns_to_split = []
        for col in df.columns:
            if df[col].astype(str).str.contains(separator, na=False).any():
                columns_to_split.append(col)
    
    # If no columns to split, return original dataframe
    if not columns_to_split:
        print("No columns found with separator to split.")
        return df
    
    print(f"Splitting columns: {columns_to_split}")
    
    # Create a copy of the dataframe
    result_df = df.copy()
    
    # Process each column that needs splitting
    for col in columns_to_split:
        # Split the column and explode into separate rows
        result_df[col] = result_df[col].astype(str).str.split(separator)
        result_df = result_df.explode(col)
    
    # Clean up: remove leading/trailing whitespace and handle empty values
    for col in columns_to_split:
        result_df[col] = result_df[col].str.strip()
        # Replace empty strings with NaN if needed
        result_df[col] = result_df[col].replace('', np.nan)
    
    # Reset index
    result_df = result_df.reset_index(drop=True)
    
    return result_df

def read_and_split_excel(file_path, sheet_name=None, columns_to_split=None, separator=';'):
    """
    Read Excel file and split specified columns containing separator into new rows
    
    Args:
        file_path (str): Path to Excel file
        sheet_name (str): Sheet name to read (None for first sheet)
        columns_to_split (list): List of column names to split
        separator (str): Character to split on
    
    Returns:
        pandas.DataFrame: Processed DataFrame
    """
    try:
        # Read Excel file
        print(f"Reading Excel file: {file_path}")
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)
        
        print(f"Original data shape: {df.shape}")
        print(f"Original columns: {list(df.columns)}")
        
        # Show sample of original data
        print("\nOriginal data (first 5 rows):")
        print(df.head())
        
        # Split columns
        split_df = split_semicolon_columns(df, columns_to_split, separator)
        
        print(f"\nProcessed data shape: {split_df.shape}")
        
        return split_df
    
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
        return None
    except Exception as e:
        print(f"Error reading Excel file: {str(e)}")
        return None

def write_split_data_to_excel(df, output_path, sheet_name='Split_Data'):
    """
    Write the processed DataFrame to Excel file
    
    Args:
        df (pandas.DataFrame): Processed DataFrame
        output_path (str): Output file path
        sheet_name (str): Sheet name for output
    """
    try:
        df.to_excel(output_path, sheet_name=sheet_name, index=False)
        print(f"Successfully wrote processed data to: {output_path}")
    except Exception as e:
        print(f"Error writing Excel file: {str(e)}")

def create_sample_data_with_semicolons():
    """
    Create sample data with semicolon-separated values for demonstration
    
    Returns:
        pandas.DataFrame: Sample data
    """
    sample_data = {
        'ID': [1, 2, 3, 4, 5],
        'Name': ['John Doe', 'Jane Smith', 'Bob Johnson', 'Alice Brown', 'Charlie Wilson'],
        'Skills': [
            'Python; Java; SQL',
            'JavaScript; HTML; CSS; React',
            'Excel; PowerPoint',
            'Python; Machine Learning; Statistics; R',
            'Project Management; Leadership'
        ],
        'Hobbies': [
            'Reading; Swimming',
            'Cooking; Traveling; Photography',
            'Gaming',
            'Hiking; Painting; Music',
            'Sports; Movies; Coding'
        ],
        'Department': ['IT', 'Web Dev', 'Admin', 'Data Science', 'Management'],
        'Salary': [75000, 65000, 45000, 85000, 95000]
    }
    
    return pd.DataFrame(sample_data)

def demonstrate_column_splitting():
    """
    Demonstrate the column splitting functionality with sample data
    """
    print("=== Excel Column Splitting Demo ===\n")
    
    # 1. Create sample data
    print("1. Creating sample data with semicolon-separated values...")
    sample_df = create_sample_data_with_semicolons()
    print(sample_df)
    print()
    
    # 2. Save sample data to Excel
    sample_file = "sample_data_with_semicolons.xlsx"
    sample_df.to_excel(sample_file, index=False)
    print(f"Sample data saved to: {sample_file}")
    print()
    
    # 3. Read and split specific columns
    print("3. Splitting 'Skills' column only:")
    split_skills_df = read_and_split_excel(
        sample_file, 
        columns_to_split=['Skills']
    )
    if split_skills_df is not None:
        print("\nProcessed data (first 10 rows):")
        print(split_skills_df.head(10))
        write_split_data_to_excel(split_skills_df, "split_skills_only.xlsx")
    print()
    
    # 4. Split multiple columns
    print("4. Splitting both 'Skills' and 'Hobbies' columns:")
    split_multiple_df = read_and_split_excel(
        sample_file, 
        columns_to_split=['Skills', 'Hobbies']
    )
    if split_multiple_df is not None:
        print("\nProcessed data (first 15 rows):")
        print(split_multiple_df.head(15))
        write_split_data_to_excel(split_multiple_df, "split_multiple_columns.xlsx")
    print()
    
    # 5. Auto-detect and split all columns with semicolons
    print("5. Auto-detecting and splitting all columns with semicolons:")
    split_auto_df = read_and_split_excel(sample_file)
    if split_auto_df is not None:
        print("\nProcessed data (first 20 rows):")
        print(split_auto_df.head(20))
        write_split_data_to_excel(split_auto_df, "split_all_semicolon_columns.xlsx")
    
    print("\n=== Demo completed! ===")
    print("Files created:")
    print("- sample_data_with_semicolons.xlsx (original data)")
    print("- split_skills_only.xlsx (Skills column split)")
    print("- split_multiple_columns.xlsx (Skills and Hobbies split)")
    print("- split_all_semicolon_columns.xlsx (all semicolon columns split)")

def advanced_split_example():
    """
    Show advanced splitting with different separators and edge cases
    """
    print("\n=== Advanced Splitting Examples ===\n")
    
    # Create data with different separators and edge cases
    advanced_data = {
        'Product_ID': ['P001', 'P002', 'P003', 'P004'],
        'Product_Name': ['Laptop', 'Phone', 'Tablet', 'Monitor'],
        'Categories': [
            'Electronics; Computers; Gaming',
            'Electronics; Mobile; Communication',
            'Electronics; Tablets',
            'Electronics; Monitors; Gaming; Office'
        ],
        'Tags': [
            'portable;lightweight;business',  # Different spacing
            'smartphone; android; 5G; ',      # Trailing semicolon and space
            'touchscreen;tablet',             # No spaces
            'large; 4K; curved; '             # Multiple trailing spaces and semicolon
        ],
        'Price': [999.99, 599.99, 399.99, 299.99]
    }
    
    advanced_df = pd.DataFrame(advanced_data)
    print("Advanced sample data:")
    print(advanced_df)
    
    # Save and process
    advanced_file = "advanced_sample.xlsx"
    advanced_df.to_excel(advanced_file, index=False)
    
    # Split with cleaning
    print("\nSplitting Categories and Tags columns:")
    split_advanced_df = read_and_split_excel(
        advanced_file, 
        columns_to_split=['Categories', 'Tags']
    )
    
    if split_advanced_df is not None:
        print("\nAdvanced processed data:")
        print(split_advanced_df)
        write_split_data_to_excel(split_advanced_df, "advanced_split_result.xlsx")

# Custom function for specific use cases
def process_excel_file(input_file, output_file, columns_to_split=None, separator=';', sheet_name=None):
    """
    Main function to process an Excel file by splitting semicolon-separated columns
    
    Args:
        input_file (str): Path to input Excel file
        output_file (str): Path to output Excel file
        columns_to_split (list): Specific columns to split (None for auto-detect)
        separator (str): Character to split on
        sheet_name (str): Sheet name to read
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Read and process the file
        processed_df = read_and_split_excel(
            input_file, 
            sheet_name=sheet_name,
            columns_to_split=columns_to_split, 
            separator=separator
        )
        
        if processed_df is not None:
            # Write the processed data
            write_split_data_to_excel(processed_df, output_file)
            return True
        else:
            return False
            
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        return False

def main():
    """
    Main function with demonstration and examples
    """
    # Run the basic demonstration
    demonstrate_column_splitting()
    
    # Run advanced examples
    advanced_split_example()
    
    print("\n=== Usage Examples ===")
    print("# Basic usage:")
    print("process_excel_file('input.xlsx', 'output.xlsx')")
    print()
    print("# Split specific columns:")
    print("process_excel_file('input.xlsx', 'output.xlsx', columns_to_split=['Skills', 'Tags'])")
    print()
    print("# Use different separator:")
    print("process_excel_file('input.xlsx', 'output.xlsx', separator='|')")
    print()
    print("# Process specific sheet:")
    print("process_excel_file('input.xlsx', 'output.xlsx', sheet_name='Sheet2')")

if __name__ == "__main__":
    main()