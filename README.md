# IFC DataToExcel

## Description

**IFC DataToExcel** is a Python application designed to facilitate the extraction of data from IFC files and its conversion into Excel format. This tool provides an intuitive graphical user interface for selecting IFC files, specifying categories of data to extract, and determining the location for saving the resulting Excel file. It leverages the `ifcopenshell` library for reading IFC files and uses `pandas` and `openpyxl` for generating and customizing the Excel output.

## Features

- User-friendly interface for selecting IFC files.
- Option to specify a save location for the Excel file.
- Ability to choose and extract specific categories of data from the IFC file.
- Automated creation of Excel sheets with extracted data.
- Customized Excel output with styled headers and adjusted column widths.

## Requirements

- Python 3.x
- `ifcopenshell`
- `pandas`
- `tkinter` (comes with standard Python installations)
- `openpyxl`

### Installation

1. Install the required Python packages:
    ```bash
    pip install ifcopenshell pandas openpyxl
    ```

## Usage

1. Run the application:
    ```bash
    ifcdata.py
    ```

2. In the GUI:
    - Click **"Browse IFC File"** to select an IFC file from your filesystem.
    - Click **"Browse Excel File"** to choose a location and name for the output Excel file.
    - Select the categories you wish to extract from the listbox.
    - Click **"Select All Categories"** to select all available categories, or **"Deselect All Categories"** to clear your selection.
    - Click **"Extract"** to start the extraction process.

3. After extraction, the data will be saved to the specified Excel file, and a confirmation message will be displayed.

## Customization

- The Excel file's appearance is customized with a green background for column headers and auto-adjusted column widths.


## Contributing

Contributions are welcome! Please submit a pull request or open an issue to report bugs or suggest features.


## Contact

For any questions or feedback, please reach out to [chiaraderosa.arch@gmail.com](mailto:chiaraderosa.arch@gmail.com).
