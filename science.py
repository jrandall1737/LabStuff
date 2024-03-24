import pandas as pd
import csv


def read_first_column_excel(excel_path):
    # Load the Excel file
    df = pd.read_excel(excel_path)

    # Assuming the first column is at position 0, extract its values
    first_column_data = df.iloc[:, 0].tolist()

    return first_column_data


def read_first_column_csv(csv_path):
    # Initialize an empty list to store the first column data
    first_column_data = []

    # Open the CSV file
    with open(csv_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)

        # Loop through each row in the CSV file
        for row in reader:
            # Add the first column of each row to the list if the row is not empty
            if row and row[0]:
                first_column_data.append(row[0])

    return first_column_data


def find_anchor_points(data, window):
    anchor_points = []
    for index, _ in enumerate(data[window:], start=window):
        if data[index] > data[index-1]:
            anchor_points.append(index)

    return anchor_points


def create_windowed_data(data, anchor_points, window):
    windowed_data = [[]]
    for anchor_point in anchor_points:
        windowed_data.append(data[anchor_point-window:anchor_point+window])
    return windowed_data


# Replace 'your_file.xlsx' with the path to your actual Excel file
excel_path = 'data.csv'
window = 6
rr_data = read_first_column_csv(excel_path)
anchor_points = find_anchor_points(rr_data, window)
windowed_data = create_windowed_data(rr_data, anchor_points, window)
