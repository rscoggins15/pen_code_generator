# reads in IOS sheet and outputs pen table
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os

# File explorer to select IOS file
root_file = tk.Tk()
root_file.withdraw()  # Hide the root window
ios_file = filedialog.askopenfilename(
    title="Select Index of Sheets File",
    initialdir=os.getcwd(),
    filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
)
root_file.destroy()

if not ios_file:
    import sys
    sys.exit()

try:
    df = pd.read_excel(ios_file, sheet_name='INDX')
    df = df[['sht_num','category','description']]
except FileNotFoundError:
    messagebox.showerror("Error", "File not found. Please check the file path.")
    import sys
    sys.exit()
except KeyError:
    messagebox.showerror("Error", "Excel file is missing required columns: 'sht_num', 'category', 'description'")
    import sys
    sys.exit()
except Exception as e:
    messagebox.showerror("Error", f"Failed to load Excel file: {str(e)}")
    import sys
    sys.exit()

def get_user_inputs():
    """Create a tkinter GUI to collect user inputs"""
    inputs = {}
    cancelled = False
    
    def submit():
        """Submit the form"""
        values = {
            'control': control_entry.get(),
            'section': section_entry.get(),
            'job': job_entry.get(),
            'hwy': hwy_entry.get(),
            'district': district_entry.get(),
            'county': county_entry.get(),
            'year': year_entry.get(),
            'funding': funding_entry.get()
        }
        
        inputs.update(values)
        inputs['cancelled'] = False
        root.quit()
    
    def cancel_form():
        """Cancel the form"""
        inputs['cancelled'] = True
        root.quit()
    
    root = tk.Tk()
    root.title("PEN Table Generator")
    root.resizable(False, False)
    
    # Set window size
    window_width = 450
    window_height = 480
    
    # Center the window on screen
    root.withdraw()  # Hide window while configuring
    root.update_idletasks()  # Update to get accurate dimensions
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    root.deiconify()  # Show window
    
    # Add menu bar
    menubar = tk.Menu(root)
    root.config(menu=menubar)
    
    help_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Help", menu=help_menu)
    help_menu.add_command(label="About", command=lambda: messagebox.showinfo("About", "PEN Table Generator v1.0\n\nGenerates PEN table files from Excel Index of Sheets."))
    
    # Add header
    header_frame = tk.Frame(root, bg="#f0f0f0", pady=10)
    header_frame.pack(fill="x")
    
    title_label = tk.Label(header_frame, text="Project Information", font=("Arial", 14, "bold"), bg="#f0f0f0")
    title_label.pack()
    
    subtitle_label = tk.Label(header_frame, text="Enter your project details", font=("Arial", 10), bg="#f0f0f0", fg="#666")
    subtitle_label.pack()
    
    # Main content frame
    content_frame = tk.Frame(root)
    content_frame.pack(fill="both", expand=True, padx=15, pady=15)
    
    # Create labels and entry fields
    tk.Label(content_frame, text="Control Number:").grid(row=0, column=0, sticky="e", padx=5, pady=8)
    control_entry = tk.Entry(content_frame, width=25)
    control_entry.grid(row=0, column=1, padx=5, pady=8)
    
    tk.Label(content_frame, text="Section Number:").grid(row=1, column=0, sticky="e", padx=5, pady=8)
    section_entry = tk.Entry(content_frame, width=25)
    section_entry.grid(row=1, column=1, padx=5, pady=8)
    
    tk.Label(content_frame, text="Job Number:").grid(row=2, column=0, sticky="e", padx=5, pady=8)
    job_entry = tk.Entry(content_frame, width=25)
    job_entry.grid(row=2, column=1, padx=5, pady=8)
    
    tk.Label(content_frame, text="Highway Name:").grid(row=3, column=0, sticky="e", padx=5, pady=8)
    hwy_entry = tk.Entry(content_frame, width=25)
    hwy_entry.grid(row=3, column=1, padx=5, pady=8)
    
    tk.Label(content_frame, text="District:").grid(row=4, column=0, sticky="e", padx=5, pady=8)
    district_entry = tk.Entry(content_frame, width=25)
    district_entry.grid(row=4, column=1, padx=5, pady=8)
    
    tk.Label(content_frame, text="County:").grid(row=5, column=0, sticky="e", padx=5, pady=8)
    county_entry = tk.Entry(content_frame, width=25)
    county_entry.grid(row=5, column=1, padx=5, pady=8)
    
    tk.Label(content_frame, text="Year:").grid(row=6, column=0, sticky="e", padx=5, pady=8)
    year_entry = tk.Entry(content_frame, width=25)
    year_entry.grid(row=6, column=1, padx=5, pady=8)
    
    tk.Label(content_frame, text="Funding:").grid(row=7, column=0, sticky="e", padx=5, pady=8)
    funding_entry = tk.Entry(content_frame, width=25)
    funding_entry.grid(row=7, column=1, padx=5, pady=8)
    
    # Button frame
    button_frame = tk.Frame(content_frame)
    button_frame.grid(row=8, column=0, columnspan=2, pady=20)
    
    # Submit and Cancel buttons
    submit_btn = tk.Button(button_frame, text="Submit", command=submit, width=12, bg="#4CAF50", fg="white", relief="raised")
    submit_btn.pack(side="left", padx=5)
    
    cancel_btn = tk.Button(button_frame, text="Cancel", command=cancel_form, width=12, bg="#f44336", fg="white", relief="raised")
    cancel_btn.pack(side="left", padx=5)
    
    # Keyboard shortcuts
    root.bind("<Return>", lambda e: submit())
    root.bind("<Escape>", lambda e: cancel_form())
    
    # Set focus to first field
    control_entry.focus()
    
    root.mainloop()
    return inputs

# Get inputs from GUI
user_inputs = get_user_inputs()

# Exit if user cancelled
if user_inputs.get('cancelled', False):
    import sys
    sys.exit()

control = user_inputs.get('control', '')
section = user_inputs.get('section', '')
job = user_inputs.get('job', '')
hwy = user_inputs.get('hwy', '')
district = user_inputs.get('district', '')
county = user_inputs.get('county', '')
year = user_inputs.get('year', '')
funding = user_inputs.get('funding', '')

# control = "0199"
# section ="01"
# job = "090"
# hwy = "US 69"
# district = "TYL"
# county = "SMITH"
# year = "2026"
# funding = "STP 2026(743)HES"

# Description abbreviations
abbreviations = {
    "SUMMARY OF SMALL SIGNS": "SOSS",
    "CONSTRUCTION SEQUENCE OF WORK": "SOW",
    "SUPPLEMENTAL INDEX OF SHEETS": "IOS",
    "QUANTITY SUMMARY SHEETS": "QS",
    "MISCELLANEOUS DETAILS": "MISC",
    "PROJECT LAYOUTS": "P",
    "TREATMENT FOR VARIOUS EDGE CONDITIONS": "EDGE_CON",
    "TYPICAL SECTIONS": "TYP",
    "GENERAL NOTES": "G_NOTES",
    "ESTIMATE AND QUANTITY SHEET": "EQ",
    "ENVIRONMENTAL PERMITS, ISSUES AND COMMITMENTS (EPIC)": "EPIC",
    "STORMWATER POLLUTION PREVENTION PLAN (SW3P)": "SW3P"
}

def get_abbreviated_description(description):
    """
    Return the abbreviated form of a description if it exists in the abbreviations mapping,
    otherwise return the original description.
    """
    return abbreviations.get(description, description)


def thru_logic(description, instance):
    """
    Parse a description with THRU pattern and generate intermediate instances.
    
    Args:
        description: String like "BC(1)-21 THRU BC(12)-21" or "TCP(1-1)-18 THRU TCP(1-5)-18"
        instance: Starting value for output numbers
    
    Returns:
        Dictionary with formatted keys and values
    """
    import re
    
    parts = description.split(" THRU ")
    if len(parts) != 2:
        return {}
    
    first_part = parts[0]  # e.g., "BC(1)-21" or "TCP(1-1)-18"
    last_part = parts[1]   # e.g., "BC(12)-21" or "TCP(1-5)-18"
    
    # Extract content in parentheses (handles both single numbers and hyphenated numbers)
    first_match = re.search(r'\(([^\)]+)\)', first_part)
    last_match = re.search(r'\(([^\)]+)\)', last_part)
    
    if not first_match or not last_match:
        return {}
    
    first_num_str = first_match.group(1)  # e.g., "1" or "1-1"
    last_num_str = last_match.group(1)    # e.g., "12" or "1-5"
    
    # Extract the last number in the sequence (handle cases like "1-1" -> extract "1")
    first_last_num = int(first_num_str.split('-')[-1])
    last_last_num = int(last_num_str.split('-')[-1])
    
    # Get the prefix and suffix
    prefix = first_part[:first_match.start()]  # e.g., "BC(" or "TCP("
    suffix = first_part[first_match.end():]    # e.g., ")-21"
    
    # Determine the base pattern (everything before the last number)
    first_base = first_num_str.rsplit('-', 1)[0] if '-' in first_num_str else ""
    
    # Generate output dictionary
    output = {}
    
    for i in range(first_last_num, last_last_num + 1):
        # Reconstruct the pattern
        if first_base:
            inner_pattern = f"{first_base}-{i}"
        else:
            inner_pattern = str(i)
        
        key = f'"${prefix}({inner_pattern}){suffix}$"'
        value = str(instance + (i - first_last_num))
        output[key] = value
    
    return output



# Write formatted output to text file
with open(f'{control}-{section}-{job}_PEN_TABLE.tbl', 'w') as f:
    # Write header
    f.write(f"""BEGIN_GLOBAL
    VERSION = 890
    PLOTTING
    VIEWS = 1-8
    SYMBOLOGY = AsStored
    EXPLODE_SHARED_CELLS = 0
    EXPLODE_DIMENSIONS = 0
    EXPLODE_MULTILINES = 0
    EXPLODE_TAGS = 1
    MATCH_MULTIPLE_SECTIONS = 0
    PST_COMPATIBLE_MODE = 0
    SORT_EXPORTED_GRAPHICS = 0
END_GLOBAL

BEGIN NEW
END

<Comments>THIS IS AN EXAMPLE TAKEN FROM AN ACTUAL PROJECT - EDIT AS NEEDED</Comments>

BEGIN_STRINGS
    "$FPN$" = "{funding}"
    "$ROADWAY NAME$" = "{hwy}"
    "$C$" = "{control}"
    "$S$" = "{section}"
    "$J$" = "{job}"
    "$HWY$" = "{hwy}"
    "$DST$" = "{district}"
    "$CTY$" = "{county}"
    "$YEAR$" = "{year}"
    "$DATE$" = "_DATE_"
    "$TIME$" = "_TIME_"
    "$FILE$" = "_FILE_"
END_STRINGS


BEGIN_STRINGS
""")
    
    counter = 0
    current_category = None
    for _, row in df.iterrows():
        category = row['category']
        # Write category header if it's new and not NaN
        if pd.notna(category) and category != current_category:
            f.write(f'\n<{category} SECTION>\n\n')
            current_category = category
        
        sht_num = row['sht_num']
        if pd.notna(sht_num) and sht_num != 0:
            description = row['description']
            # Apply abbreviations
            description = get_abbreviated_description(description)
            
            # Add blank line before if more than 1 sheet
            # if int(sht_num) > 1:
            #     f.write('\n')
            
            if "THRU" in description:
                # Use thru_logic to generate all instances
                thru_dict = thru_logic(description, counter + 1)
                for key, value in thru_dict.items():
                    line = f'\t{key} = "{value}"'
                    f.write(line + '\n')
                counter += len(thru_dict)
            else:
                for i in range(int(sht_num)):
                    counter += 1
                    instance = i + 1
                    if int(sht_num) > 1:
                        line = f'\t"${description}-{instance}$" = "{int(counter)}"'
                    else:
                        line = f'\t"${description}$" = "{int(counter)}"'
                    f.write(line + '\n')
            
            # Add blank line after if more than 1 sheet
            if int(sht_num) > 1:
                f.write('\n')
    
    # Write footer
    f.write('END_STRINGS\n')

