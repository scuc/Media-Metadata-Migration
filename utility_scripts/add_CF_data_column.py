import pandas as pd
import argparse
import os


def add_cf_output_dir_column(input_file, output_file=None):
    """
    Add a new CF_OUTPUT_DIR column to the Excel file based on rules for TitleType and Content_Type.

    Updated Rules:
    - { TitleType: Video } = Video
    - { TitleType: Document } = Documents
    - { TitleType: graphic } = Packages
    - {TitleType: archive} and {Content_Type: [AVP, FCP, PPRO, PTS]} = Projects
    - {TitleType: archive} and {Content_Type contains WAV} = Audio (now checks if WAV is in the string)
    - {TitleType: archive} and {Content_Type: GRFX} = Packages

    Args:
        input_file (str): Path to the input Excel file
        output_file (str, optional): Path to save the output Excel file with the new column

    Returns:
        pd.DataFrame: DataFrame with the new CF_OUTPUT_DIR column
    """
    # Read the Excel file
    print(f"Reading Excel file: {input_file}")
    df = pd.read_excel(input_file)

    # Check if necessary columns exist
    if "TITLETYPE" not in df.columns:
        print("Warning: TITLETYPE column not found in the Excel file.")
        return df

    if "CONTENT_TYPE" not in df.columns:
        print("Warning: CONTENT_TYPE column not found in the Excel file.")
        return df

    # Define the project content types (case insensitive)
    project_content_types = ["AVP", "FCP", "PPRO", "PTS"]

    # Define a function to apply the rules for CF_OUTPUT_DIR
    def determine_output_dir(row):
        # Convert values to appropriate case and handle null/undefined
        title_type = str(row["TITLETYPE"]).lower() if pd.notna(row["TITLETYPE"]) else ""
        content_type = (
            str(row["CONTENT_TYPE"]).upper() if pd.notna(row["CONTENT_TYPE"]) else ""
        )

        # Apply the primary rules in order
        if title_type == "video":
            return "video"
        elif title_type == "document":
            return "document"
        elif title_type == "graphic":
            return "package"
        elif title_type == "archive":
            # Exact matches first
            if content_type in project_content_types:
                return "project"
            elif content_type == "GRFX":
                return "package"
            # Then partial matches
            elif "WAV" in content_type:
                return "audio"
            # Check for partial matches in project content types
            elif any(proj_type in content_type for proj_type in project_content_types):
                return "project"
            elif "GRFX" in content_type:
                return "package"

        # Additional flexible rule checks for ANY title_type (not just archive)
        # This catches files that might have been miscategorized
        if any(proj_type in content_type for proj_type in project_content_types):
            return "project"
        elif "WAV" in content_type or "AUDIO" in content_type:
            return "audio"
        elif "VIDEO" in content_type or "MOV" in content_type or "MP4" in content_type:
            return "video"
        elif "DOC" in content_type or "PDF" in content_type or "TXT" in content_type:
            return "document"
        elif (
            "GRFX" in content_type
            or "PNG" in content_type
            or "JPG" in content_type
            or "TIFF" in content_type
        ):
            return "package"

        # Default value if no rules match
        return "unknown"

    # Apply the function to create the new column
    df["CF_OUTPUT_DIR"] = df.apply(determine_output_dir, axis=1)

    # Count the occurrences of each value in CF_OUTPUT_DIR
    value_counts = df["CF_OUTPUT_DIR"].value_counts()
    print("\nCF_OUTPUT_DIR value counts:")
    for value, count in value_counts.items():
        print(f"  - {value}: {count}")

    # Sample rows with different CF_OUTPUT_DIR values (if any)
    unique_output_dirs = df["CF_OUTPUT_DIR"].unique()
    if len(unique_output_dirs) > 1:
        print("\nSample rows for each unique CF_OUTPUT_DIR value:")
        for output_dir in unique_output_dirs:
            sample_rows = df[df["CF_OUTPUT_DIR"] == output_dir]
            if not sample_rows.empty:
                sample_row = sample_rows.iloc[0]
                print(f"\nExample for CF_OUTPUT_DIR = '{output_dir}':")
                print(f"  TitleType: {sample_row['TITLETYPE']}")
                print(f"  Content_Type: {sample_row['CONTENT_TYPE']}")
                if "NAME" in df.columns:
                    print(f"  NAME: {sample_row['NAME']}")
            else:
                print(f"\nNo examples found for CF_OUTPUT_DIR = '{output_dir}'")
    else:
        print(f"\nAll rows have the same CF_OUTPUT_DIR value: {unique_output_dirs[0]}")

    # If the "unknown" category is still large, print some samples to help identify patterns
    if "unknown" in value_counts and value_counts["unknown"] > 10:
        unknown_samples = df[df["CF_OUTPUT_DIR"] == "unknown"].head(5)
        print("\nSample of 'unknown' categorized rows for further investigation:")
        for idx, row in unknown_samples.iterrows():
            print(f"  Sample {idx}:")
            print(f"    TitleType: {row['TITLETYPE']}")
            print(f"    Content_Type: {row['CONTENT_TYPE']}")
            if "NAME" in df.columns:
                print(f"    NAME: {row['NAME']}")

    # Save to output file if specified
    if output_file:
        print(f"\nSaving updated data to: {output_file}")
        df.to_excel(output_file, index=False)
        print(f"File saved successfully!")

    return df


def main():
    # Set up argument parser
    parser = argparse.ArgumentParser(
        description="Add CF_OUTPUT_DIR column to an Excel file based on rules."
    )
    parser.add_argument("input_file", help="Path to the input Excel file")
    parser.add_argument(
        "-o",
        "--output",
        default="output_with_cf_dir.xlsx",
        help="Path to save the output Excel file (default: output_with_cf_dir.xlsx)",
    )

    args = parser.parse_args()

    # Add CF_OUTPUT_DIR column
    updated_df = add_cf_output_dir_column(args.input_file, args.output)

    # Print summary
    print(f"\nAdded CF_OUTPUT_DIR column to the Excel file.")
    print(f"The new column has been determined based on the following rules:")
    print("  - TitleType: Video → CF_OUTPUT_DIR = video")
    print("  - TitleType: Document → CF_OUTPUT_DIR = document")
    print("  - TitleType: graphic → CF_OUTPUT_DIR = package")
    print(
        "  - TitleType: archive & Content_Type in [AVP, FCP, PPRO, PTS] → CF_OUTPUT_DIR = project"
    )
    print("  - TitleType: archive & Content_Type contains WAV → CF_OUTPUT_DIR = audio")
    print("  - TitleType: archive & Content_Type = GRFX → CF_OUTPUT_DIR = package")
    print("\nAdditional flexible rules applied to reduce 'unknown' values:")
    print(
        "  - Content_Type contains any of [AVP, FCP, PPRO, PTS] → CF_OUTPUT_DIR = project"
    )
    print("  - Content_Type contains WAV or AUDIO → CF_OUTPUT_DIR = audio")
    print("  - Content_Type contains VIDEO, MOV, MP4 → CF_OUTPUT_DIR = video")
    print("  - Content_Type contains DOC, PDF, TXT → CF_OUTPUT_DIR = document")
    print("  - Content_Type contains GRFX, PNG, JPG, TIFF → CF_OUTPUT_DIR = package")


if __name__ == "__main__":
    main()
