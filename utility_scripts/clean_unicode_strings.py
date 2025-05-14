import pandas as pd
import re
import os
import argparse


def clean_filename(filename):
    """
    Clean filenames by removing non-standard character sequences like '‚Äã'.

    Args:
        filename (str): The original filename with non-standard characters

    Returns:
        str: The cleaned filename
    """
    if not filename or not isinstance(filename, str):
        return filename

    # Remove the '‚Äã' sequence
    cleaned = re.sub(r"‚Äã", "", filename)

    # Fix hyphen issue found in some filenames
    cleaned = re.sub(r"‚Äã-‚Äã", "-", cleaned)

    # Handle any other non-standard characters
    cleaned = re.sub(
        r"[^\x00-\x7F]", "", cleaned
    )  # Remove any remaining non-ASCII characters

    return cleaned


def process_excel_file(
    input_file, output_file=None, rename_actual_files=False, file_directory=None
):
    """
    Process an Excel file containing filenames with non-standard characters.

    Args:
        input_file (str): Path to the input Excel file
        output_file (str, optional): Path to save the output Excel file with cleaned names
        rename_actual_files (bool): Whether to rename actual files on disk
        file_directory (str, optional): Directory containing the files to rename

    Returns:
        pd.DataFrame: DataFrame with cleaned filenames
    """
    # Read the Excel file
    print(f"Reading Excel file: {input_file}")
    df = pd.read_excel(input_file)

    # Check if the necessary columns exist
    name_cols = []
    for col in ["NAME", "FILENAME", "OBJECTNM"]:
        if col in df.columns:
            name_cols.append(col)

    if not name_cols:
        raise ValueError("No filename columns found in the Excel file.")

    # Create a new DataFrame for output with all original columns
    output_df = df.copy()

    # Clean the filename columns and replace them in the output DataFrame
    for col in name_cols:
        # Store the original values for reporting
        original_values = output_df[col].copy()

        # Replace the values with cleaned versions
        output_df[col] = output_df[col].apply(clean_filename)

        # Print some statistics
        changed_count = sum(original_values != output_df[col])
        print(f"Cleaned {changed_count} out of {len(df)} values in {col} column")

    # Save the cleaned data to a new Excel file (with all columns but cleaned filenames)
    if output_file:
        print(f"Saving cleaned data to: {output_file}")
        output_df.to_excel(output_file, index=False)

    # Rename actual files if requested
    if rename_actual_files and file_directory:
        if "FILENAME" not in name_cols:
            print("WARNING: Cannot rename files because FILENAME column doesn't exist.")
        else:
            renamed_count = 0
            for i, row in df.iterrows():
                old_name = os.path.join(file_directory, row["FILENAME"])
                new_name = os.path.join(file_directory, output_df.iloc[i]["FILENAME"])

                if os.path.exists(old_name) and old_name != new_name:
                    try:
                        os.rename(old_name, new_name)
                        renamed_count += 1
                    except Exception as e:
                        print(f"Error renaming {old_name}: {e}")

            print(f"Renamed {renamed_count} files on disk.")

    # For reporting purposes, create a DataFrame with both original and cleaned names
    report_df = df.copy()
    for col in name_cols:
        report_df[f"CLEANED_{col}"] = output_df[col]

    return report_df


def main():
    # Set up argument parser
    parser = argparse.ArgumentParser(
        description="Clean filenames with non-standard characters in an Excel file."
    )
    parser.add_argument("input_file", help="Path to the input Excel file")
    parser.add_argument(
        "-o",
        "--output",
        default="cleaned_names.xlsx",
        help="Path to save the output Excel file with cleaned names (default: cleaned_names.xlsx)",
    )
    parser.add_argument(
        "-r", "--rename", action="store_true", help="Rename actual files on disk"
    )
    parser.add_argument(
        "-d", "--directory", help="Directory containing the files to rename"
    )
    parser.add_argument(
        "-p",
        "--preview",
        action="store_true",
        help="Preview changes without saving a new file",
    )

    args = parser.parse_args()

    # Process the Excel file
    output_file = None if args.preview else args.output

    report_df = process_excel_file(
        args.input_file, output_file, args.rename, args.directory
    )

    # Print some examples of changes
    if not report_df.empty:
        print("\nExamples of cleaned filenames:")
        examples_shown = 0

        for i in range(min(len(report_df), 20)):  # Check first 20 rows for examples
            for col in ["NAME", "FILENAME", "OBJECTNM"]:
                if col in report_df.columns and f"CLEANED_{col}" in report_df.columns:
                    orig = report_df.iloc[i][col]
                    cleaned = report_df.iloc[i][f"CLEANED_{col}"]
                    if orig != cleaned:  # Only show examples that changed
                        print(f"Original {col}: {orig}")
                        print(f"Cleaned {col}: {cleaned}")
                        print("---")
                        examples_shown += 1
                        if examples_shown >= 5:  # Show at most 5 examples
                            break
            if examples_shown >= 5:
                break

    if not args.preview and output_file:
        print(f"\nCleaned Excel file saved to: {output_file}")
        print(
            "The file contains the same columns and rows as the original, but with cleaned filenames."
        )

    return report_df


if __name__ == "__main__":
    main()
