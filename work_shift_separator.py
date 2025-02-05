import pandas as pd

def split_shifts(input_file, output_file):
    """
    Split Excel data into two shifts while preserving all row information.
    
    Args:
        input_file (str): Path to input Excel file
        output_file (str): Path to output Excel file
    """
    # Read the Excel file
    df = pd.read_excel(input_file)
    
    # Initialize lists for shift1 and shift2
    shift1_records = []
    shift2_records = []
    
    # Keep track of previous row for handling shift2
    prev_row = None
    
    for idx, row in df.iterrows():
        # If person_group is empty and we have a previous row, this is a shift2 record
        if pd.isna(row['person_group']) and prev_row is not None:
            # Create shift2 record with number and person_group from previous row
            shift2_row = row.copy()
            shift2_row['number'] = prev_row['number']
            shift2_row['person_group'] = prev_row['person_group']
            shift2_records.append(shift2_row)
        else:
            # This is a shift1 record
            shift1_records.append(row)
            prev_row = row
    
    # Create DataFrames
    shift1_df = pd.DataFrame(shift1_records)
    shift2_df = pd.DataFrame(shift2_records)
    
    # Write to Excel
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        shift1_df.to_excel(writer, sheet_name='Shift 1', index=False)
        if not shift2_df.empty:
            shift2_df.to_excel(writer, sheet_name='Shift 2', index=False)
            
    # Print summary
    print(f"Shift 1 records: {len(shift1_df)}")
    print(f"Shift 2 records: {len(shift2_df)}")

# Example usage
if __name__ == "__main__":
    split_shifts('input.xlsx', 'output.xlsx')