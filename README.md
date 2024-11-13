# Data Transformation and Cleaning Automation Tool

This project is a powerful **Data Transformation and Cleaning Automation Tool** developed in Python to automate the ETL (Extract, Transform, Load) process. It includes data cleaning, conditional formatting, and report generation with charts in an Excel format, designed to facilitate efficient data analysis and presentation.

## Features

- **Data Extraction**: Loads data from a source file (`.xlsx`) and prepares it for transformation.
- **Data Transformation and Cleaning**: Normalizes columns, removes duplicates, fills missing values, and applies conditional formatting to highlight values below a specified threshold.
- **Report Generation**:
  - **Pie Chart**: Visualizes the distribution of a specified categorical column (e.g., `category`).
- **Excel Formatting**: Provides readable, professional data presentations through conditional formatting and embedded charts in Excel.

## Project Structure

```bash
Data-Transformation-and-Cleaning-Automation-Tool/
├── data/                        # Folder for storing raw data files
├── reports/                     # Folder to save generated Excel reports
├── scripts/
│   ├── extract_data.py          # Script to load data from source files
│   ├── transform_data.py        # Script to clean and transform data
│   ├── load_data.py             # Script to save processed data
│   ├── report_generator.py      # Script to generate formatted Excel report
├── main.py                      # Main script to run the ETL pipeline
└── README.md                    # Project documentation
```

## Setup and Installation

### Prerequisites

- **Python 3.7+**
- **Pandas** and **OpenPyXL** for data processing and Excel handling
- **NumPy** for statistical calculations

### Installation

1. **Clone the Repository**:

   ```bash
   git clone https://github.com/your-username/Data-Transformation-and-Cleaning-Automation-Tool.git
   cd Data-Transformation-and-Cleaning-Automation-Tool
   ```

2. **Create a Virtual Environment**:

   ```bash
   python3 -m venv venv
   source venv/bin/activate   # On Windows, use venv\Scripts\activate
   ```

3. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

### Running the Project

1. Place your input data file (`.xlsx` format) in the `data/` directory, ensuring it has the necessary columns (e.g., `signup_date`, `purchase_amount`, `category`).
2. Run the ETL pipeline:
   ```bash
   python main.py
   ```
3. The output report will be saved in the `reports/` directory as `report.xlsx`.

## Key Features and Functionality

### 1. Data Transformation and Cleaning

- **Column Normalization**: Converts column names to a standardized format.
- **Duplicate Removal**: Ensures each data entry is unique.
- **Missing Values Handling**: Fills missing values in key columns (`age`, `score`) with the column mean.

### 2. Conditional Formatting

- **Threshold-Based Highlighting**: Applies red color fill to values below a specified threshold, making data anomalies easily visible.

### 3. Excel Report Generation

- **Pie Chart**: Shows the distribution of a specified categorical column, e.g., `category`, with labels displaying category names, values, and percentages.

## Sample Output

The generated Excel report (`report.xlsx`) includes:

- **Conditional Formatting**: Highlighted cells based on specified criteria for quick identification of outliers.
- **Pie Chart**: A pie chart showing the distribution of a categorical column, making it easier to interpret categorical data.

## Skills Demonstrated

- **Data Transformation and Automation**: Automates ETL processes with Python, handling tasks like data extraction, transformation, and validation.
- **Data Visualization**: Creates engaging and informative charts directly in Excel using `openpyxl`.
- **Data Cleaning**: Implements data integrity checks, including handling duplicates, missing values, and outliers.

## Future Enhancements

- **Email Automation**: Send the generated report as an email attachment.
- **Database Integration**: Enable data extraction from SQL databases for more flexible data sources.
- **Dynamic Charting**: Allow more customization for charts based on user-defined parameters.
