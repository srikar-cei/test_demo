import json
import pandas as pd

def json_to_excel(json_file_path, excel_file_path):
    # Load the JSON data
    with open(json_file_path, 'r') as file:
        data = json.load(file)

    # Extract relevant information for Excel
    steps = data.get('steps', [])
    excel_data = []

    for step in steps:
        excel_data.append({
            "Step": step['step'],
            "Description": step['description'],
            "Expected Result": step['expected_result'],
            "Actual Result": step['actual_result'],
            "Status": step['status']
        })

    # Create a DataFrame
    df = pd.DataFrame(excel_data)

    # Write DataFrame to Excel
    df.to_excel(excel_file_path, index=False, sheet_name='Test Results')

    print(f"Excel report saved to '{excel_file_path}'")

# Example usage
json_file_path = 'final_test_report.json'  # Path to your JSON file
excel_file_path = 'og_test_report.xlsx'  # Desired path for the Excel file
json_to_excel(json_file_path, excel_file_path)
