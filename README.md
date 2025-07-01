OVERVIEW:

## SRA Automated Extract-N-Fill Tool

This tool provides a graphical interface for extracting property owner information 
from various Texas county CAD (Central Appraisal District) websites 
based on property ID input from an Excel spreadsheet. 
Data is scraped using Selenium and saved to a new Excel file with 
the property owner’s name and address to their assigned IDs.


TECHNOLOGIES USED:  - Python 3.x
- Tkinter (GUI)
- Selenium WebDriver (Chrome)
- OpenPyXL (Excel handling)


FEATURES:

- Select a county from a dropdown
- Choose Excel file containing property IDs
- Set output directory
- Automate browser to scrape property info
- Save results to Excel


HOW TO RUN:

(If you are running this program through the executable, skip to step 5)

1. Clone or download this repo.
2. Ensure `chromedriver` is installed and matches your Chrome version.
3. Install required Python packages:
pip install selenium openpyxl
4. Run the script through the executable file or run the GUI class on an IDE or code editor.

(For general users, double-click the “SRA CAD Extract-N-Fill” executable file.)

5. Provided in the folder is an Excel file called “test.xlsx”. This can be used to input the Property IDs. 
Excel Spreadsheet Columns:
	•	Column A = prop_id (These must be already inputted by the user
	•	Column B = owner_name
	•	Column C = owner_address
6. Select a county, upload an Excel file, and run the scraper.
7. A clone Excel file will be generated with the complete filled information.


GUI FLOW DOCUMENTATION:  

### GUI Workflow

1. User selects a county from dropdown
2. User uploads an Excel file with property IDs (column A)
3. User selects an output folder
4. User clicks "Run"
5. The program:
   - Loads input IDs
   - Launches the selected county's script
   - Uses Selenium to scrape data
   - Writes data to a new Excel file in the output folder
Error Handling & Limitations

### Error Handling

- Empty or invalid property IDs are skipped
- If no clickable result appears for an ID, the program logs it and continues
- WebDriver is safely closed after the operation

### Known Limitations

- Requires stable internet connection
- Website structure may change (XPath may need updates)













