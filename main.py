import argparse
import csv
import datetime
import pandas as pd
import os
import sys

import dc_calculation


def validate_directory(directory):
    # Check if the directory exists
    if os.path.isdir(directory) == False:
        print(f"The provided directory {directory} is not a valid directory.")
        sys.exit(1)


def get_excel_file_paths(directory):
    # Check if the directory contains an excel file
    excel_files = [file for file in os.listdir(
        directory) if file.endswith('.xlsx') or file.endswith('.xls')]
    if len(excel_files) == 0:
        print(f"No excel files (.xlsx or .xls) files found in {directory}.")
        sys.exit(1)

    file_paths = []
    for file in excel_files:
        file_paths.append(os.path.abspath(os.path.join(directory, file)))

    return file_paths


def create_rr_data_dictionary(excel_path):
    # Replace 'your_file.xlsx' with the path to your Excel file
    column_name = 'RR-I(ms):ECG'
    print(f'Reading file {excel_path}, this can take a while...')

    # Read the Excel file
    # df = pd.read_excel(excel_path)
    # Read all sheets from an excel file
    excel_file_sheets = pd.read_excel(excel_path, sheet_name=None)

    data_dict = dict()

    # Select the column
    for sheet_name, df in excel_file_sheets.items():
        column_data = df[column_name]
        if column_data.empty:
            print(
                f'Column with name {column_name} not found in sheet {sheet_name}')
            exit(1)

        numerical_data = column_data.apply(pd.to_numeric, errors='coerce')
        numerical_data = numerical_data.dropna()
        dict_name = get_file_name(excel_path) + "_" + sheet_name
        data_dict[dict_name] = numerical_data

    return data_dict


def write_results_file(dc_dictionary, directory):
    # create a timestamped filename
    now = datetime.datetime.now()
    results_file_path = os.path.join(
        directory, f'dc_results-{now.strftime("%Y-%m-%d_%H%M%S")}.csv')

    print(f'Creating combined results file {results_file_path}')

    # Open a file for writing
    with open(results_file_path, 'w', newline='') as csvfile:
        # Create a CSV writer object
        writer = csv.writer(csvfile)

        for key, value in dc_dictionary.items():
            writer.writerow([key, value])


def get_file_name(file_path):
    # get the name of the data file without the path or file extension
    data_file_name, _ = os.path.splitext(os.path.basename(file_path))
    return data_file_name


def main():
    # Create the parser
    parser = argparse.ArgumentParser(description='Process a directory name.')

    # Add an argument for the directory
    parser.add_argument('--directory',
                        type=str,
                        help='The name of the directory where the collected data is stored in excel spreadsheets.',
                        default='.',
                        required=False)

    # Parse the command line arguments
    args = parser.parse_args()

    # check that the path is a directory and contains excel files
    validate_directory(args.directory)
    data_files_paths = get_excel_file_paths(args.directory)
    print(data_files_paths)

    dc_dictionary = dict()

    for data_file_path in data_files_paths:
        rr_data_dict = create_rr_data_dictionary(data_file_path)

        for data_name, rr_data_unfiltered in rr_data_dict.items():
            dc = dc_calculation.run_analysis(
                rr_data_unfiltered, data_name, 12, 0.05)
            dc_dictionary[data_name] = dc

    write_results_file(dc_dictionary, args.directory)


if __name__ == "__main__":
    main()
