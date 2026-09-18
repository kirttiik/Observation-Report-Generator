import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import math
import os

def to_int(val):
    try:
        if pd.isna(val) or str(val).strip() == '' or val is None:
            return None
        return int(float(val))
    except (ValueError, TypeError):
        return None

def clean_school_name(name):
    """
    Cleans the name so it can be used as a valid Excel sheet name.
    Excel restricts sheet names to 31 characters and no special characters.
    """
    name = str(name).strip()
    invalid_chars = [':', '\\', '/', '?', '*', '[', ']']
    for ch in invalid_chars:
        name = name.replace(ch, '')
    # Truncate to 31 chars max
    return name[:31]

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Teacher Observation Report Generator")
        self.root.geometry("650x550")
        self.root.configure(padx=20, pady=20)
        
        # Variables
        self.file1_var = tk.StringVar()
        self.file2_var = tk.StringVar()
        self.file3_var = tk.StringVar()
        self.out_var = tk.StringVar()
        
        self.col_school_var = tk.StringVar(value="1")     # A
        self.col_teachers_var = tk.StringVar(value="3")   # C
        self.col_unique_var = tk.StringVar(value="6")     # F (Corrected to 6, as Image 1 shows F. 4 was D: Target)
        self.col_obs_var = tk.StringVar(value="5")        # E (Rule 3 says 5th)
        self.col_term_obs_var = tk.StringVar(value="8")   # 8th column for file 2 and 3
        
        self.create_widgets()

    def create_widgets(self):
        style = ttk.Style()
        style.configure('TButton', font=('Calibri', 10))
        style.configure('TLabel', font=('Calibri', 10))
        
        # --- File Selection Section ---
        frame_files = ttk.LabelFrame(self.root, text="Select Excel Files", padding=(10, 10))
        frame_files.pack(fill='x', pady=10)
        
        # File 1
        ttk.Label(frame_files, text="File 1 (Entire Year: Apr '25 - Apr '26):").grid(row=0, column=0, sticky='w', pady=5)
        ttk.Entry(frame_files, textvariable=self.file1_var, width=50).grid(row=0, column=1, padx=10, pady=5)
        ttk.Button(frame_files, text="Browse", command=lambda: self.browse_file(self.file1_var)).grid(row=0, column=2, pady=5)
        
        # File 2
        ttk.Label(frame_files, text="File 2 (Term 1: Apr '25 - Oct '25):").grid(row=1, column=0, sticky='w', pady=5)
        ttk.Entry(frame_files, textvariable=self.file2_var, width=50).grid(row=1, column=1, padx=10, pady=5)
        ttk.Button(frame_files, text="Browse", command=lambda: self.browse_file(self.file2_var)).grid(row=1, column=2, pady=5)
        
        # File 3
        ttk.Label(frame_files, text="File 3 (Term 2: Oct '25 - Apr '26):").grid(row=2, column=0, sticky='w', pady=5)
        ttk.Entry(frame_files, textvariable=self.file3_var, width=50).grid(row=2, column=1, padx=10, pady=5)
        ttk.Button(frame_files, text="Browse", command=lambda: self.browse_file(self.file3_var)).grid(row=2, column=2, pady=5)
        
        # Output File
        ttk.Label(frame_files, text="Output Save Location:").grid(row=3, column=0, sticky='w', pady=5)
        ttk.Entry(frame_files, textvariable=self.out_var, width=50).grid(row=3, column=1, padx=10, pady=5)
        ttk.Button(frame_files, text="Browse", command=self.browse_save_file).grid(row=3, column=2, pady=5)
        
        # --- Config Section ---
        frame_config = ttk.LabelFrame(self.root, text="Column Configuration (1-based index)", padding=(10, 10))
        frame_config.pack(fill='x', pady=10)
        ttk.Label(frame_config, text="School Name Col:").grid(row=0, column=0, sticky='w', pady=2)
        ttk.Entry(frame_config, textvariable=self.col_school_var, width=10).grid(row=0, column=1, padx=5, pady=2, sticky='w')
        
        ttk.Label(frame_config, text="No. of Teachers Col:").grid(row=0, column=2, sticky='w', pady=2)
        ttk.Entry(frame_config, textvariable=self.col_teachers_var, width=10).grid(row=0, column=3, padx=5, pady=2, sticky='w')
        
        ttk.Label(frame_config, text="Unique Teacher Col:").grid(row=1, column=0, sticky='w', pady=2)
        ttk.Entry(frame_config, textvariable=self.col_unique_var, width=10).grid(row=1, column=1, padx=5, pady=2, sticky='w')
        ttk.Label(frame_config, foreground='gray', text="(Your Rule 5 stated 4th column, adjust to 6 if it's F)").grid(row=1, column=2, columnspan=2, sticky='w', padx=5)
        
        ttk.Label(frame_config, text="Observation Till Date Col:").grid(row=2, column=0, sticky='w', pady=2)
        ttk.Entry(frame_config, textvariable=self.col_obs_var, width=10).grid(row=2, column=1, padx=5, pady=2, sticky='w')
        ttk.Label(frame_config, foreground='gray', text="(Your Rule 3 stated 5th column)").grid(row=2, column=2, columnspan=2, sticky='w', padx=5)

        ttk.Label(frame_config, text="Term 1 & 2 Obs Col:").grid(row=3, column=0, sticky='w', pady=2)
        ttk.Entry(frame_config, textvariable=self.col_term_obs_var, width=10).grid(row=3, column=1, padx=5, pady=2, sticky='w')
        ttk.Label(frame_config, foreground='gray', text="(8th column as per latest rule)").grid(row=3, column=2, columnspan=2, sticky='w', padx=5)

        # --- Action Section ---
        frame_action = ttk.Frame(self.root)
        frame_action.pack(fill='x', pady=20)
        
        btn_generate = ttk.Button(frame_action, text="Generate Reports", command=self.process_files)
        btn_generate.pack(ipadx=20, ipady=10)
        
    def browse_file(self, var):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if filename:
            var.set(filename)

    def browse_save_file(self):
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if filename:
            self.out_var.set(filename)

    def extract_data(self, file_path, col_school, col_obs, req_teacher_and_unique=False, col_teachers=0, col_unique=0):
        try:
            try:
                # 1. Try standard read (for .xlsx)
                df = pd.read_excel(file_path, header=None)
                dfs_to_process = [df]
            except ValueError:
                # 2. If it complains about engine, it's likely a standard .xls
                try:
                    df = pd.read_excel(file_path, header=None, engine='xlrd')
                    dfs_to_process = [df]
                except Exception:
                    # 3. If it still fails, it might be an HTML file masquerading as .xls
                    dfs_to_process = pd.read_html(file_path, header=None)
        except Exception as e:
            raise Exception(f"Failed to read {file_path}:\n{str(e)}")

        data = {}
        for df_item in dfs_to_process:
            for idx, row in df_item.iterrows():
                try:
                    # Ensure the dataframe has enough columns
                    if len(row) <= max(col_school, col_obs, col_teachers, col_unique):
                        continue

                    school = str(row.iloc[col_school]).strip()
                    school_key = school.upper()
                    
                    # Skip meaningless rows or header rows
                    if not school or pd.isna(row.iloc[col_school]) or school.lower() in ['nan', 'school', 'session: 2025-2026', 'staff observation report']:
                        continue

                    obs_int = to_int(row.iloc[col_obs])
                    if obs_int is None:
                        continue

                    if school_key not in data:
                        item = {'obs': obs_int, 'name': school}
                        if req_teacher_and_unique:
                            teachers_int = to_int(row.iloc[col_teachers])
                            unique_int = to_int(row.iloc[col_unique])
                            if teachers_int is None or unique_int is None:
                                continue
                            item['teachers'] = teachers_int
                            item['unique'] = unique_int
                        data[school_key] = item
                except IndexError:
                    continue
                except Exception as e:
                    continue

        return data

    def process_files(self):
        f1 = self.file1_var.get()
        f2 = self.file2_var.get()
        f3 = self.file3_var.get()
        out = self.out_var.get()

        if not all([f1, f2, f3, out]):
            messagebox.showerror("Missing Information", "Please select all three input files and an output save location.")
            return

        try:
            col_school = int(self.col_school_var.get()) - 1
            col_teachers = int(self.col_teachers_var.get()) - 1
            col_unique = int(self.col_unique_var.get()) - 1
            col_obs = int(self.col_obs_var.get()) - 1
            col_term_obs = int(self.col_term_obs_var.get()) - 1
        except ValueError:
            messagebox.showerror("Invalid Input", "Column indexes must be numbers.")
            return

        try:
            # Extract
            data1 = self.extract_data(f1, col_school, col_obs, req_teacher_and_unique=True, col_teachers=col_teachers, col_unique=col_unique)
            data2 = self.extract_data(f2, col_school, col_term_obs)
            data3 = self.extract_data(f3, col_school, col_term_obs)

            if not data1:
                messagebox.showerror("Error", "No valid school data found in File 1 based on the selected columns.")
                return

            # Note: We need openpyxl installed to write
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                for school_key, d1 in data1.items():
                    school_name = d1.get('name', school_key)
                    teachers = d1.get('teachers', 0)
                    obs_all = d1.get('obs', 0)
                    unique = d1.get('unique', 0)

                    # Rule 2: Teachers * 8
                    target = teachers * 8
                    
                    # Rule 4: % till date
                    perc_all = round((obs_all / target * 100), 2) if target > 0 else 0
                    
                    # Rule 6: Teachers not observed once (can't be negative)
                    not_observed = max(0, teachers - unique)

                    # --- Term 1 ---
                    d2 = data2.get(school_key, {})
                    # Rule 7: Target / 2
                    term1_target = int(target / 2)
                    term1_obs = d2.get('obs', 0)
                    # Rule 9: % Term 1
                    perc_1 = round((term1_obs / term1_target * 100), 2) if term1_target > 0 else 0
                    
                    # --- Term 2 ---
                    d3 = data3.get(school_key, {})
                    # Rule 10: Target / 2
                    term2_target = int(target / 2)
                    term2_obs = d3.get('obs', 0)
                    # Rule 11: % Term 2
                    perc_2 = round((term2_obs / term2_target * 100), 2) if term2_target > 0 else 0

                    output_data = [
                        ['No. of Teachers', teachers],
                        ['No. of Target observations (8 per teacher per year x no of teacher)', target],
                        ['Observations till date (1st/April/2025 to 1st/April/2026)', obs_all],
                        ['To observation Percentage till date', perc_all],
                        ['Unique Teacher Count', unique],
                        ['No. of Teachers not observed once', not_observed],
                        ['', ''],
                        ['Term Target 1 (1/April/2025 to 15/oct/2025)', term1_target],
                        ['Observation Term 1', term1_obs],
                        ['Percentage of Term 1 observation', perc_1],
                        ['', ''],
                        ['Term Target 2 (16/Oct/2025 to 1st/April/2026)', term2_target],
                        ['Observation Term 2', term2_obs],
                        ['Percentage of Term 2 observation', perc_2]
                    ]

                    df_out = pd.DataFrame(output_data, columns=['Observation', 'Count'])
                    sname = clean_school_name(school_name)
                    
                    # Write sheet
                    df_out.to_excel(writer, sheet_name=sname, index=False)
                    
                    # Format sheet to look like new format picture (Red header, wide columns)
                    worksheet = writer.sheets[sname]
                    
                    # Optional: Import openpyxl styles inside the loop to avoid missing module errors before reading data
                    from openpyxl.styles import PatternFill, Font

                    header_fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
                    header_font = Font(color="FFFFFF", bold=True)

                    for cell in worksheet["1:1"]:
                        cell.fill = header_fill
                        cell.font = header_font
                        
                    # Adjust column widths
                    worksheet.column_dimensions['A'].width = 65
                    worksheet.column_dimensions['B'].width = 15

            messagebox.showinfo("Success", f"Report generated successfully!\nSaved to: {out}")
            
        except Exception as e:
            messagebox.showerror("Processing Error", f"An error occurred while processing:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
