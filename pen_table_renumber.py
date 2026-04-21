import re
import os
import tkinter as tk
from tkinter import filedialog

# File explorer to select pen table file
root_file = tk.Tk()
root_file.withdraw()  # Hide the root window
pen_tbl = filedialog.askopenfilename(
    title="Select Pen Table",
    initialdir=os.getcwd(),
    filetypes=[("Pen Table", "*.tbl"), ("All files", "*.*")]
)
root_file.destroy()

# Exit if no file was selected
if not pen_tbl:
    import sys
    sys.exit()

def renumber_after_general_section(
    input_path: str,
    output_path: str,
    start_number: int = 1
):
    """
    Renumbers quoted numeric values sequentially,
    but ONLY after the <GENERAL SECTION> marker.
    """

    value_pattern = re.compile(r'(=\s*")([^"]*)(")')
    current_number = start_number
    output_lines = []

    renumbering_active = False

    with open(input_path, "r", encoding="utf-8") as infile:
        for line in infile:

            # Detect the boundary where renumbering begins
            if "<GENERAL SECTION>" in line or "<Comments>GENERAL SECTION</Comments>" in line:
                renumbering_active = True
                output_lines.append(line)
                continue

            # Before GENERAL SECTION: copy lines unchanged
            if not renumbering_active:
                output_lines.append(line)
                continue

            # After GENERAL SECTION: renumber numeric values
            def replace(match):
                nonlocal current_number
                value = match.group(2)

                if value.isdigit():
                    new_value = str(current_number)
                    current_number += 1
                    return f'{match.group(1)}{new_value}{match.group(3)}'

                return match.group(0)

            new_line = value_pattern.sub(replace, line)
            output_lines.append(new_line)

    with open(output_path, "w", encoding="utf-8") as outfile:
        outfile.writelines(output_lines)


# =========================
# SCRIPT ENTRY POINT
# =========================
if __name__ == "__main__":
    renumber_after_general_section(
        input_path=pen_tbl,
        output_path=pen_tbl.replace(".tbl", "-1.tbl"),
        start_number=1
    )
