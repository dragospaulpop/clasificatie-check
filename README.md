# XLSX Data Comparison Tool

This Node.js application scans, analyzes, and compares Excel data from various files against a main dictionary. It is designed to identify new data entries which are present in the file but not found in the main dictionary. The resulting output is an xlsx file named 'missing-codes.xlsx', containing these missing data entries.

## Setup Instructions

1. **File Placement**:
    - Position the xlsx files to be compared inside the `/files` folder.
    - Position the main classification xlsx file (the reference dataset) in the root folder. This file should be named 'clasificatie.xlsx'.

2. **Dependency Installation**:
    - Run the following command to install the required dependencies:
        ```
        npm install
        ```

3. **Execution**:
    - With the setup complete, execute the application using the command:
        ```
        node index.js
        ```

## Code Overview

This application utilizes the 'xlsx' and 'fs-extra' npm packages to read the excel files and perform operations.

The code involves a series of functions designed to carry out specific tasks:

- `xlsxToJson`: This function reads an xlsx file and returns a JSON representation of the data.
- `buildClasificatieDictionary`: This function constructs a dictionary from the main classification file (clasificatie.xlsx).
- `extractCodesFromDataFile`: This function extracts unique codes from a given file and returns a dictionary of these codes.
- `main`: This is the primary function. It coordinates the operations of reading files, building dictionaries, comparing data and writing the output to an xlsx file.

Upon execution, the script scans the provided files, compares them with the main dictionary, and generates an output file, 'missing-codes.xlsx', containing the missing data entries.

With this XLSX Data Comparison Tool, you can automate the comparison of multiple Excel files, saving time and effort by identifying any missing data entries.
