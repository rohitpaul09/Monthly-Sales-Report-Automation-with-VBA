# Monthly Sales Report Automation with VBA

## Project Description

This VBA project is designed to automate various tasks associated with monthly sales reports in Excel. The project includes several subroutines that work together to streamline the process of creating pivot tables, generating charts, copying Excel content to PowerPoint for presentations, and cleaning up the workbook by deleting pivot tables and charts. This automation is particularly useful for businesses or individuals regularly dealing with sales data and requiring consistent report formats.

## Features

### 1. `CreatePivot` Subroutine
- **Functionality**: Automatically creates pivot tables across all worksheets (except "MacroButtons") in the workbook.
- **Key Operations**: 
  - Deletes existing pivot tables.
  - Generates new pivot tables with specific fields for Branch, Date, Category, and Revenue.
  - Formats data fields.

### 2. `CreateChart` Subroutine
- **Functionality**: Generates column clustered charts for each worksheet (excluding "MacroButtons").
- **Key Operations**: 
  - Deletes any existing charts.
  - Creates new charts and sets their properties, including source data range and title.

### 3. `Copy_Excel_To_PPT` Subroutine
- **Functionality**: Copies charts from Excel sheets to a new PowerPoint presentation.
- **Key Operations**: 
  - Creates a new PowerPoint slide for each chart.
  - Formats the slide with a title and adjusts the chart’s size and position.

### 4. `DeletePivotTablesAndCharts` Subroutine
- **Functionality**: Cleans up worksheets by removing all pivot tables and charts.
- **Key Operations**: 
  - Iterates through each worksheet to clear pivot tables and delete chart objects.

## Installation

To use these subroutines:
1. Open your Excel workbook.
2. Press `Alt + F11` to open the VBA editor.
3. Insert a new module (Insert > Module).
4. Copy and paste the provided VBA code into the module.
5. Close the VBA editor and return to your Excel workbook.

## Usage

Run the desired subroutine based on your needs:
- Use `CreatePivot` to generate pivot tables for data analysis.
- Use `CreateChart` to visualize data with charts.
- Use `Copy_Excel_To_PPT` to prepare PowerPoint presentations with your data.
- Use `DeletePivotTablesAndCharts` to clean up the workbook before starting a new report cycle.

## A dedicated sheet named "MacroButtons" contains buttons for each subroutine:

- Create Pivot Tables: Runs CreatePivot.
- Create Charts: Executes CreateChart.
- PowerPoint Presentation: Initiates Copy_Excel_To_PPT.
- Clear Charts & Pivots: Activates DeletePivotTablesAndCharts.
Click the respective button to run a subroutine.

## Customization

You can customize these subroutines by modifying parameters and properties in the VBA code to fit your specific data structure and reporting needs.

## Compatibility

This project is compatible with Microsoft Excel. Ensure your version of Excel supports VBA.

## Contribution

Contributions to improve or enhance the functionality of these subroutines are welcome. Please submit pull requests with a detailed description of your changes.

## Contact

For queries or suggestions, contact me at rohitpaul09@gmail.com.


