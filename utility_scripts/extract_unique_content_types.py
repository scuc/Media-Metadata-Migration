import pandas as pd
import argparse
import re
import os


def extract_unique_content_types(
    input_file, output_file=None, column_name="CONTENT_TYPE"
):
    """
    Extract unique content types from a specified column in the Excel file.
    Handles comma-separated values and space-separated values.

    Args:
        input_file (str): Path to the input Excel file
        output_file (str, optional): Path to save the output text file with unique content types
        column_name (str): Name of the column to extract values from (default: 'CONTENT_TYPE')

    Returns:
        list: List of unique content types
    """
    # Read the Excel file
    print(f"Reading Excel file: {input_file}")
    df = pd.read_excel(input_file)

    # Check if the specified column exists
    if column_name not in df.columns:
        print(f"Warning: {column_name} column not found in the Excel file.")
        return []

    # Extract all content types
    all_content_types = []

    # Drop null values and convert to string
    content_type_values = df[column_name].dropna().astype(str)

    print(f"Sample values from {column_name} column:")
    for value in content_type_values.head(5):  # Show first 5 values as samples
        print(f'  - "{value}"')

    # Process each value in the column
    for value in content_type_values:
        # Skip empty values
        if not value or value.lower() == "nan":
            continue

        # First try splitting by comma
        if "," in value:
            parts = [part.strip() for part in value.split(",")]
        else:
            # If no commas, try splitting by spaces
            parts = [part.strip() for part in re.split(r"\s+", value)]

        # Add non-empty items to the list
        all_content_types.extend([part for part in parts if part])

    # Get unique values and sort them
    unique_content_types = sorted(list(set(all_content_types)))

    # Print results
    print(f"\nFound {len(unique_content_types)} unique values in {column_name}:")
    for content_type in unique_content_types:
        print(f"  - {content_type}")

    # Save to output file if specified
    if output_file:
        with open(output_file, "w") as f:
            f.write(f"Unique {column_name} Values:\n")
            for content_type in unique_content_types:
                f.write(f"{content_type}\n")
        print(f"Saved unique values to: {output_file}")

    return unique_content_types


def main():
    # Set up argument parser
    parser = argparse.ArgumentParser(
        description="Extract unique values from a column in an Excel file."
    )
    parser.add_argument("input_file", help="Path to the input Excel file")
    parser.add_argument(
        "-c",
        "--column",
        default="CONTENT_TYPE",
        help="Column name to extract values from (default: CONTENT_TYPE)",
    )
    parser.add_argument(
        "-o", "--output", help="Path to save the output text file with unique values"
    )

    args = parser.parse_args()

    # Extract unique values
    unique_values = extract_unique_content_types(
        args.input_file, args.output, args.column
    )

    # Print summary
    print(
        f"\nSummary: Found {len(unique_values)} unique values in the {args.column} column."
    )
    print("As a comma-separated list:")
    print(f"{args.column} = [{','.join(unique_values)}]")


if __name__ == "__main__":
    main()
