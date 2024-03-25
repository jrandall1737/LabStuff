import argparse
import csv
import datetime
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


def write_results_file(dc_dictionary, directory):
    # create a timestamped filename
    now = datetime.datetime.now()
    results_file_path = os.path.join(
        directory, f'dc_results-{now.strftime("%Y-%m-%d_%H%M%S")}.csv')

    # Open a file for writing
    with open(results_file_path, 'w', newline='') as csvfile:
        # Create a CSV writer object
        writer = csv.writer(csvfile)

        for key, value in dc_dictionary.items():
            writer.writerow([key, value])


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
        dc = dc_calculation.run_analysis(data_file_path, 12, 0.05)
        dc_dictionary[os.path.basename(data_file_path)] = dc

    write_results_file(dc_dictionary, args.directory)


if __name__ == "__main__":
    main()
