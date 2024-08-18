import pandas as pd
from tkinter import Tk, filedialog, messagebox, Button, Text, Scrollbar, END
import os
import glob
from tkinter import END

def excel_to_cfg(excel_file, cfg_file):
    df = pd.read_excel(excel_file)
    
    with open(cfg_file, 'w') as file:
        for _, row in df.iterrows():
            section = row[0].replace("(String)", "").strip()
            file.write(f'[{section}]\n')
            
            for col_name, value in row[1:].items():
                if isinstance(value, bool):
                    value_str = str(value).lower()
                elif "(String)" in col_name:
                    col_name_clean = col_name.replace("(String)", "").strip()
                    value_str = f'"{value}"'
                else:
                    value_str = value
                
                file.write(f'{col_name_clean if "(String)" in col_name else col_name} = {value_str}\n')
            
            file.write('\n')


def convert_xlsx_to_csv(output_widget):
    excel_file = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if not excel_file:
        output_widget.insert(END, "No file selected!\n")
        return

    csv_file = filedialog.asksaveasfilename(
        title="Save CSV File",
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv")]
    )

    if not csv_file:
        output_widget.insert(END, "No save location selected!\n")
        return
    
    df = pd.read_excel(excel_file)
    df.to_csv(csv_file, index=False, sep=',', encoding='utf-8')
    output_widget.insert(END, f"CSV file saved to {csv_file}\n")

def manual_convert(output_widget):
    excel_file = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if not excel_file:
        output_widget.insert(END, "No file selected!\n")
        return

    cfg_file = filedialog.asksaveasfilename(
        title="Save CFG File",
        defaultextension=".cfg",
        filetypes=[("CFG files", "*.cfg")]
    )

    if not cfg_file:
        output_widget.insert(END, "No save location selected!\n")
        return

    excel_to_cfg(excel_file, cfg_file)
    output_widget.insert(END, f"CFG file saved to {cfg_file}\n")

def main():
    root = Tk()
    root.title("Excel Converter")
    root.geometry("500x400")

    # Create a Text widget with a scrollbar
    output_widget = Text(root, wrap="word", height=10)
    output_widget.pack(pady=10, padx=10, expand=True, fill="both")
    
    scrollbar = Scrollbar(output_widget)
    scrollbar.pack(side="right", fill="y")
    output_widget.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=output_widget.yview)

    Button(root, text="Convert XLSX to CFG", command=lambda: manual_convert(output_widget), width=25).pack(pady=5)
    Button(root, text="Convert XLSX to CSV", command=lambda: convert_xlsx_to_csv(output_widget), width=25).pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    main()
