import pandas as pd
import zipfile
import argparse
import os


def parse_arguments():
    parser = argparse.ArgumentParser()

    parser.add_argument("--zip_file", type=str, help="Name of the downloaded zip file.")
    parser.add_argument("--unziped_dir", type=str, help="Directory to unzip the file.")
    parser.add_argument("--xlsx", type=str, help="Name of the XLSX file to process.")
    parser.add_argument("--category", type=str, help="Name of the category to filter.")
    parser.add_argument("--output_file", type=str, help="Name of the output CSV file")

    return parser.parse_args()


def unzip_file(file: str, dir: str):
    with zipfile.ZipFile(file, "r") as zip_ref:
        zip_ref.extractall(dir)


def process_csv(file: str, dir: str, category: str):
    path = os.path.join(dir, file)

    if os.path.isfile(path):
        df = pd.read_excel(path, header=None)

        # Remove header and footer
        # I asume the header is always the same large (5 rows), otherwise I could search for the cell with 'Total'
        header = df.head(5)
        df = df.iloc[6:]
        if df.iloc[-1].count() == 1:
            df = df.iloc[:-1]
        df = df.reset_index(drop=True)

        df = filter_category(df, category)

        return pd.concat([header, df])
    else:
        print(f"{file} was not found in {dir}.")
        return None


def filter_category(df: pd.DataFrame, category: str):
    category_indexes = df[df[0].str.startswith("Crimes Against")].index
    if len(category_indexes) == 0:
        return pd.DataFrame()

    for i, index in enumerate(category_indexes):
        if i < len(category_indexes) - 1:
            if df[0].iloc[index] == category:
                return df.iloc[index : category_indexes[i + 1]]
        else:
            if df[0].iloc[index] == category:
                return df.iloc[index:]


if __name__ == "__main__":
    args = parse_arguments()

    unzip_file(args.zip_file, args.unziped_dir)

    df = process_csv(args.xlsx, args.unziped_dir, args.category)
    df.to_csv(f"{args.output_file}.csv", index=False)
