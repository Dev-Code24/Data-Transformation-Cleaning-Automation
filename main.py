from scripts.scripts import *

def main():
    # Define the path for input and output files in .xlsx format
    input_file = "data/test_raw_data.xlsx"
    processed_file = "data/processed_data.xlsx"
    report_file = "reports/report.xlsx"
    
    raw_data = extract_data(input_file, file_type='excel')
    
    transformed_data = transform_data(raw_data)
    
    load_data(transformed_data, processed_file)
    
    generate_report(transformed_data, report_file)

    print("Report has been successfully generated and saved to 'reports/report.xlsx'. Exiting script.")

if __name__ == "__main__":
    main()
