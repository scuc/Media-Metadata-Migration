import pandas as pd
import math
import os


def split_csv(file_path, output_dir, chunk_size):
    """
    Splits a large CSV file into multiple smaller CSV files with evenly divided rows.

    Parameters:
    - file_path: str, path to the large CSV file
    - output_dir: str, directory where the split CSV files will be saved
    - chunk_size: int, the number of rows per split file
    """

    # Ensure the output directory exists
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Read the CSV in chunks
    csv_data = pd.read_csv(file_path, chunksize=chunk_size)

    date = os.path.basename(file_path).split("_")[0]

    # Save each chunk as a separate CSV file
    for i, chunk in enumerate(csv_data):
        output_file = os.path.join(output_dir, f"{date}_chunk_{i + 1}.csv")
        chunk.to_csv(output_file, index=False)
        print(f"Saved {output_file}")


if __name__ == "__main__":
    split_csv(
        file_path="data_CSV/202410141224_gor_diva_merged_cleaned_final.csv",
        output_dir="data_chunked",
        chunk_size=10000,
    )
