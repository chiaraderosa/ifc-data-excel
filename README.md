# IFC DataToExcel

## Description

**IFC DataToExcel** is a Python application designed to facilitate the extraction of data from IFC (Industry Foundation Classes) files and its conversion into Excel format. This tool provides an intuitive graphical user interface (GUI) for selecting IFC files, specifying categories of data to extract, and determining the location for saving the resulting Excel file. It leverages the `ifcopenshell` library for reading IFC files and uses `pandas` and `openpyxl` for generating and customizing the Excel output.

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
- `tkinter` (included with standard Python installations)
- `openpyxl`

### Installation

1. Clone the repository:
    ```bash
    git clone https://github.com/yourusername/ifc-datatoexcel.git
    cd ifc-datatoexcel
    ```

2. Install the required Python packages:
    ```bash
    pip install ifcopenshell pandas openpyxl
    ```

## Usage

1. Run the application:
    ```bash
    python main.py
    ```

2. In the GUI:
    - Click **"Browse IFC File"** to select the IFC file from your filesystem.
    - Click **"Browse Excel File"** to choose the destination and filename for the Excel output.
    - Select the categories you want to extract from the listbox.
    - Use **"Select All Categories"** to choose all available categories or **"Deselect All Categories"** to clear your selections.
    - Click **"Extract"** to begin the data extraction process.

3. After extraction, the data will be saved in the chosen Excel file, and a confirmation message will be displayed.

## Customization

- The Excel output is customized with a green background for column headers and auto-adjusted column widths for optimal readability.

## Troubleshooting

- **IFC File Loading Issues**: Ensure that the IFC file is correctly formatted and compatible with `ifcopenshell`.
- **No Categories Displayed**: The IFC file may not contain any data categories or products.

## Contributing

Contributions are welcome! To contribute, please submit a pull request or open an issue to report bugs or suggest new features.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contact

For questions or feedback, please contact [your-email@example.com](mailto:your-email@example.com).