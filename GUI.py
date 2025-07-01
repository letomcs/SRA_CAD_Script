from tkinter import *
from tkinter import filedialog, ttk
import importlib
import os

window = Tk()
window.geometry("450x520")
window.title('SRA Automated Extract-N-Fill')

# --- COUNTY SELECTION ---
Label(window, text="Select a County:").pack(anchor='w', padx=10, pady=(20, 0))

# All County options
counties = ["Newton", "Sabine", "Shelby", "Panola", "Wood", "Rains", "Hunt", "VanZandt", "Kaufman", "Hopkins"]

# Use Combobox widget for dropdown list
county_dropdown = ttk.Combobox(window, text="Select a County", values=counties, state="readonly")
county_dropdown.pack(padx=10, pady=5, fill=X)

# --- SELECT EXCEL FILE ---
Label(window, text="Open Excel Spreadsheet (.xlsx file):").pack(anchor='w', padx=10, pady=(20, 0))
input_pathname = Entry(window, width=50)
input_pathname.pack(padx=10, pady=5)

Button(window, text="Open", command=lambda: openExcelFile()).pack(anchor='w', padx=10, pady=2)

# --- OUTPUT DIRECTORY ---
Label(window, text="Select File Location to Store:").pack(anchor='w', padx=10, pady=(20, 0))
output_directory = Entry(window, width=50)
output_directory.pack(padx=10, pady=5)

Button(window, text="Select", command=lambda: selectOutputDirectory()).pack(anchor='w', padx=10, pady=2)

# --- STATUS LABEL ---
status_label = Label(window, text="", fg="black")
status_label.pack(padx=10, pady=(10, 0))

# --- FUNCTIONS ---

# Open Excel filepath
def openExcelFile():
    filepath = filedialog.askopenfilename(filetypes=[('Excel Workbook', '*.xlsx')])
    if filepath:
        input_pathname.delete(0, END)
        input_pathname.insert(0, filepath)

# Select the dirctory to store file
def selectOutputDirectory():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        output_directory.delete(0, END)
        output_directory.insert(0, folder_selected)

# Input Excel, Select county, and Select output directory 
def submit():
    input_excel = input_pathname.get()
    selected_county = county_dropdown.get()
    output_dir = output_directory.get()

    output_path = os.path.join(output_dir, f"{selected_county} Owners 1 Mile.xlsx")
    
    try:
        module = importlib.import_module(selected_county)  
        script_instance = getattr(module, selected_county)()      
        result = script_instance.process_data(input_excel, output_path)

        if result == "Success":
            status_label.config(text="✅ Excel spreadsheet filled", fg="green")
        else:
            status_label.config("A problem has occurred", fg="red")
    except Exception as e:
        status_label.config(text=f"❌ Error: {e}", fg="red")

# --- RUN BUTTON ---
Button(window, text="Run", command=submit).pack(side=BOTTOM, pady=20)

window.mainloop()
