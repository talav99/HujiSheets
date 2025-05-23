# Python Spreadsheet Application

A spreadsheet application built with **Python** and **Tkinter**, supporting formula calculations, cell formatting, and multiple file export formats.

---

##  Features
- **Spreadsheet Creation**
  - Create spreadsheets up to 50 rows × 25 columns
  - Easily enter and navigate data
- **Formula Support**  
  Use formulas with an `=` prefix:
  - Basic math: `=5+10*2`
  - Cell references: `=A1+B2`
  - Built-in functions: `=SUM(A1:A5)`, `=AVERAGE(A1:A5)`, `=MAX(...)`, `=MIN(...)`
  - Apply colors: `=clr(A1:B3)` prompts for color selection

- **Cell Formatting**
  - Change cell background and font color
  - Adjust font size
  - Color entire ranges via formulas

- **File Operations**
  - **Import**: JSON, YAML
  - **Export**: JSON, YAML, PDF, CSV, Excel (XLSX)

***Project Structure***

```plaintext
python-spreadsheet/
├── main.py              # Main application file with the user interface
├── requirements.txt     # Project dependencies
└── utils/               # Utility package
    ├── __init__.py
    ├── calculates.py    # Formula parsing and calculation logic
    ├── import_export.py # File import and export functionality
    ├── project_testing.py # Testing module
    └── sheet.py         # Core Sheet and Cell classes
```
## Installation
Installation

Clone the repository:
git clone https://github.com/yourusername/python-spreadsheet.git
cd python-spreadsheet

Create and activate a virtual environment:
# Windows
python -m venv venv
venv\Scripts\activate

# macOS/Linux
python3 -m venv venv
source venv/bin/activate

Install the required dependencies:
pip install -r requirements.txt

## Usage
**Running the Application**

Run the main script to start the application:
python main.py
You can also run with the help flag to see usage information:
python main.py --help

**Working with Cells**

Click on any cell to select it and enter a value
Press Enter to move to the cell below
Enter formulas starting with = (e.g., =A1+B2)

**Formatting**

**Use the Edit menu to:**

* Change cell background color
* Change font color
* Increase or decrease font size
* Use the =clr(A1:B3) formula to apply a color to a range of cells

## Saving and Exporting

**Use the File menu to:**

Import from JSON/YAML
Export to JSON
Export to YAML
Export to PDF
Export to CSV
Export to Excel

Formula Examples

Basic math: =5+10*2
Cell references: =A1+B2
Functions: =SUM(A1:A5)
Functions: =AVERAGE(A1:A5)
Functions: =MAX(A1:A5)
Functions: =MIN(A1:A5)
Color range: =clr(A1:B3) (will prompt for color selection)

## Dependencies
The application requires the following Python packages:

tkinter (included in standard Python)
Additional dependencies listed in requirements.txt

## License
MIT License