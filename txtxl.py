import os
import sys
import argparse
import glob
import pandas as pd
import time


def split_file_by_instruction_code(input_file, delimiter="~"):
    """
    Split a large input file into multiple files based on the Instruction Code.

    Parameters:
    - input_file: Path to the input file to be split
    - delimiter: Delimiter used in the input file

    Returns:
    - Path to the output folder containing split files
    """
    start_time = time.time()

    # Get the directory of the input file
    input_dir = os.path.dirname(os.path.abspath(input_file))

    # Create intermediate output folder in the same directory as input file
    output_folder = os.path.join(input_dir, "intermediate")
    os.makedirs(output_folder, exist_ok=True)

    # Open and process the input file
    with open(input_file, "r") as file:
        lines = file.readlines()

    # Extract header and data lines
    header = lines[0].strip()
    data_lines = lines[1:]

    # Create a dictionary to group lines by the 2nd column (Instruction Code)
    grouped_data = {}
    for line in data_lines:
        columns = line.strip().split(delimiter)
        instruction_code = columns[1]
        if instruction_code not in grouped_data:
            grouped_data[instruction_code] = []
        grouped_data[instruction_code].append(line)

    # Write each group to a separate file
    for instruction_code, records in grouped_data.items():
        output_file = os.path.join(output_folder, f"{instruction_code}.txt")
        with open(output_file, "w") as out_file:
            out_file.write(header + "\n")
            out_file.writelines(records)

    # Calculate and print processing time
    end_time = time.time()
    processing_time = end_time - start_time
    print(f"Splitting took {processing_time:.2f} seconds")
    print(f"Files created for each Instruction Code in: {output_folder}")
    
    return output_folder


def convert_text_to_excel(
    input_folder, output_folder=None, delimiter="~", max_rows=1048576
):
    """
    Convert text files to Excel, splitting large files if necessary.

    Parameters:
    - input_folder: Path to the folder containing text files
    - output_folder: Path to save Excel files (defaults to input folder)
    - delimiter: Delimiter used in text files
    - max_rows: Maximum rows per Excel file

    Returns:
    - List of created Excel file paths
    """
    start_time = time.time()

    # If no output folder specified, use input folder
    if output_folder is None:
        output_folder = input_folder

    # Create output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    # Find all text files in the input folder
    text_files = glob.glob(os.path.join(input_folder, "*.txt"))

    # List to store created Excel file paths
    excel_files = []

    # Process each text file
    for text_file in text_files:
        try:
            # Read the text file
            df = pd.read_csv(text_file, sep=delimiter, dtype=str)

            # Get base filename without extension
            base_filename = os.path.splitext(os.path.basename(text_file))[0]

            # Split file if it exceeds max rows
            if len(df) > max_rows:
                # Calculate number of split files needed
                num_splits = (len(df) + max_rows - 1) // max_rows

                for i in range(num_splits):
                    # Slice the dataframe for this split
                    start_idx = i * max_rows
                    end_idx = min((i + 1) * max_rows, len(df))
                    split_df = df.iloc[start_idx:end_idx].reset_index(drop=True)

                    # Create filename for split
                    split_filename = f"{base_filename}_part{i+1}.xlsx"
                    output_path = os.path.join(output_folder, split_filename)

                    # Save to Excel using XlsxWriter engine with large file support
                    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
                        split_df.to_excel(writer, index=False, sheet_name="Sheet1")

                    excel_files.append(output_path)
                    print(f"Created: {split_filename}")
            else:
                # If file is within row limit, save as single Excel file
                output_path = os.path.join(output_folder, f"{base_filename}.xlsx")

                # Save to Excel using XlsxWriter engine
                with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
                    df.to_excel(writer, index=False, sheet_name="Sheet1")

                excel_files.append(output_path)
                print(f"Created: {os.path.basename(output_path)}")

        except Exception as e:
            print(f"Error processing {text_file}: {e}")

    # Calculate and print processing time
    end_time = time.time()
    processing_time = end_time - start_time
    print(f"Converting to Excel took {processing_time:.2f} seconds")

    return excel_files


def main():
    # Set up argument parser
    parser = argparse.ArgumentParser(description="File Splitting and Conversion Pipeline")
    
    # Add arguments
    parser.add_argument('input_file', help='Path to the input file to be processed')
    parser.add_argument(
        '-d', '--delimiter', 
        default='~', 
        help='Delimiter used in the file (default: ~)'
    )
    parser.add_argument(
        '-m', '--max_rows', 
        type=int, 
        default=1048576, 
        help='Maximum rows per Excel file (default: 1048576)'
    )
    parser.add_argument(
        '-o', '--output_folder', 
        help='Path to save Excel files (defaults to input file directory)'
    )
    
    # Parse arguments
    args = parser.parse_args()
    
    # Get the directory of the input file
    input_dir = os.path.dirname(os.path.abspath(args.input_file))
    
    # If no output folder specified, use input file's directory
    output_folder = args.output_folder or os.path.join(input_dir, "excel_output")
    
    # Record total start time
    total_start_time = time.time()
    
    # Step 1: Split the input file by instruction code
    intermediate_folder = split_file_by_instruction_code(
        args.input_file, 
        args.delimiter
    )
    
    # Step 2: Convert split text files to Excel
    created_files = convert_text_to_excel(
        intermediate_folder, 
        output_folder, 
        args.delimiter, 
        args.max_rows
    )
    
    # Calculate and print total processing time
    total_end_time = time.time()
    total_processing_time = total_end_time - total_start_time
    
    # Print summary
    print("\nProcessing complete.")
    print(f"Total processing time: {total_processing_time:.2f} seconds")
    print(f"Total Excel files created: {len(created_files)}")
    print("Created files:")
    for file in created_files:
        print(file)


if __name__ == "__main__":
    main()