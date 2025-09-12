import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
from tkinter.scrolledtext import ScrolledText
import os
import re
from datetime import datetime
from pathlib import Path

class FormGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PM Form Generator v1.3a")
        self.root.geometry("1000x700")
        
        # Initialize variables
        self.source_file = None
        self.sheets_data = {}
        self.selected_sheet = None
        self.procedures = []
        self.sections = []
        self.user_name = ""
        self.lov_database = self.init_lov_database()
        
        self.create_widgets()
    
    def init_lov_database(self):
        """Initialize LOV database with existing codes"""
        return {
            'YKN-CPP2-G-603-DG': ['Dirty', 'Good'],
            'YKN-CPP2-G-603-CNBG': ['Clogged', 'Not Working', 'Broken', 'Good'],
            'YKN-CPP2-G-603-LG': ['Leak', 'Good'],
            'YKN-CPP2-G-603-MG': ['Miss Timing', 'Good'],
            'YKN-CPP2-G-603-DLCG': ['Dirty', 'Leak', 'Clogged', 'Good'],
            'YKN-CPP2-G-603-DCG': ['Dirty', 'Clogged', 'Good'],
            'YKN-CPP2-G-603-OG': ['Out of Spec', 'Good'],
            'YKN-CPP2-G-603-OBG': ['Offset', 'Blurred', 'Good'],
            'YKN-CPP2-G-603-DWG': ['Dirty', 'Wet', 'Good'],
            'YKN-CPP2-G-603-YN': ['Yes', 'No'],
            'YKN-CPP2-G-603-GFB': ['Good', 'Fair', 'Bad'],
            'YKN-CPP2-G-603-RR1': ['Reset', 'Replaced'],
            'YKN-CPP2-G-603-R1': ['Replaced'],
            'YKN-CPP2-G-603-RR2': ['Repaired', 'Replaced'],
            'YKN-CPP2-G-603-RC1': ['Replaced', 'Cleaned Up'],
            'YKN-CPP2-G-603-A1': ['Adjust'],
            'YKN-CPP2-G-603-CRR1': ['Cleaned up', 'Replaced', 'Retighten'],
            'YKN-CPP2-G-603-RA1': ['Replaced', 'Adjust'],
            'YKN-CPP2-G-603-CD1': ['Cleaned Up', 'Drying'],
            'YKN-CPP2-G-603-RR3': ['Replaced', 'Refill'],
            'YKN-CPP2-G-603-AR1': ['Adjust', 'Repaired'],
            'YKN-CPP2-G-603-RRC1': ['Repaired', 'Replaced', 'Cleaned Up'],
        }
    
    def create_widgets(self):
        """Create main GUI components"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="5")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Button(file_frame, text="Select Excel File", 
                  command=self.select_file).grid(row=0, column=0, padx=(0, 10))
        
        self.file_label = ttk.Label(file_frame, text="No file selected")
        self.file_label.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        # Sheet selection section
        self.sheet_frame = ttk.LabelFrame(main_frame, text="Sheet Selection", padding="5")
        self.sheet_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        self.sheet_frame.columnconfigure(1, weight=1)
        
        ttk.Label(self.sheet_frame, text="Select Tasklist Sheet:").grid(row=0, column=0, padx=(0, 10))
        
        self.sheet_combo = ttk.Combobox(self.sheet_frame, state="readonly")
        self.sheet_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        self.sheet_combo.bind('<<ComboboxSelected>>', self.on_sheet_selected)
        
        ttk.Button(self.sheet_frame, text="Analyze Sheet", 
                  command=self.analyze_sheet).grid(row=0, column=2)
        
        # User info section
        user_frame = ttk.LabelFrame(main_frame, text="User Information", padding="5")
        user_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        user_frame.columnconfigure(1, weight=1)
        
        ttk.Label(user_frame, text="Username:").grid(row=0, column=0, padx=(0, 10))
        self.username_entry = ttk.Entry(user_frame)
        self.username_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        # Analysis results section
        self.analysis_frame = ttk.LabelFrame(main_frame, text="Analysis Results", padding="5")
        self.analysis_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        self.analysis_frame.columnconfigure(0, weight=1)
        self.analysis_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        self.analysis_text = ScrolledText(self.analysis_frame, height=15, width=80)
        self.analysis_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Generate button section
        generate_frame = ttk.Frame(main_frame)
        generate_frame.grid(row=4, column=0, columnspan=2, pady=(10, 0))
        
        self.generate_btn = ttk.Button(generate_frame, text="Configure & Generate Forms", 
                                     command=self.configure_generation, state="disabled")
        self.generate_btn.pack()
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.progress.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
    def select_file(self):
        """Select source Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            self.source_file = file_path
            self.file_label.config(text=os.path.basename(file_path))
            self.load_sheets()
    
    def load_sheets(self):
        """Load sheets from Excel file"""
        try:
            # Read all sheet names
            excel_file = pd.ExcelFile(self.source_file)
            sheet_names = excel_file.sheet_names
            
            # Filter sheets that might be tasklists
            tasklist_sheets = [name for name in sheet_names 
                             if any(keyword.lower() in name.lower() 
                                   for keyword in ['task', 'list', 'maintenance', 'check'])]
            
            if not tasklist_sheets:
                tasklist_sheets = sheet_names  # Show all if no obvious tasklists
            
            self.sheet_combo['values'] = tasklist_sheets
            self.analysis_text.delete(1.0, tk.END)
            self.analysis_text.insert(tk.END, f"Loaded {len(sheet_names)} sheets.\n")
            self.analysis_text.insert(tk.END, f"Potential tasklist sheets: {len(tasklist_sheets)}\n\n")
            
            if tasklist_sheets:
                self.sheet_combo.set(tasklist_sheets[0])
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel file: {str(e)}")
    
    def on_sheet_selected(self, event):
        """Handle sheet selection"""
        self.selected_sheet = self.sheet_combo.get()
        self.generate_btn.config(state="disabled")
    
    def analyze_sheet(self):
        """Analyze selected sheet structure"""
        if not self.selected_sheet:
            messagebox.showwarning("Warning", "Please select a sheet first.")
            return
        
        if not self.username_entry.get().strip():
            messagebox.showwarning("Warning", "Please enter username first.")
            return
        
        self.user_name = self.username_entry.get().strip()
        
        try:
            # Read the sheet with openpyxl to get merged cell info
            from openpyxl import load_workbook
            wb = load_workbook(self.source_file)
            ws = wb[self.selected_sheet]
            
            # Read with pandas for data processing
            df = pd.read_excel(self.source_file, sheet_name=self.selected_sheet, header=None)
            
            # Find procedures and sections
            self.procedures = []
            self.sections = []
            current_section = ""
            last_proc_num = 0
            
            for idx, row in df.iterrows():
                # Convert row to list and handle NaN values
                row_values = [str(cell).strip() if not pd.isna(cell) else '' for cell in row]
                
                # Check for procedures in table format (number in col A, description in col B)
                if len(row_values) > 1:
                    col_a = row_values[0]  # First column (usually number)
                    col_b = row_values[1]  # Second column (usually procedure)
                    
                    # Check if col A contains a number and col B has procedure text
                    if col_a.isdigit() and len(col_b) > 5:  # Reasonable procedure length
                        proc_num = int(col_a)
                        proc_desc = col_b
                        
                        # Detect section change when numbering resets to 1
                        if proc_num == 1 and last_proc_num > 1:
                            # Look for section header above this procedure
                            section_found = False
                            # Check a few rows above for section header
                            for check_idx in range(max(0, idx - 5), idx):
                                if check_idx < len(df):
                                    check_row = df.iloc[check_idx]
                                    for cell_val in check_row:
                                        if not pd.isna(cell_val):
                                            cell_str = str(cell_val).strip()
                                            # Potential section header criteria:
                                            # - Not empty, not too long, not a procedure description
                                            if (len(cell_str) > 3 and len(cell_str) < 50 and 
                                                not cell_str.isdigit() and
                                                not any(word in cell_str.lower() for word in ['inspect', 'replace', 'check', 'test', 'measure', 'clean'])):
                                                current_section = cell_str
                                                if current_section not in [s['name'] for s in self.sections]:
                                                    self.sections.append({
                                                        'name': current_section,
                                                        'row': check_idx,
                                                        'col': 0
                                                    })
                                                section_found = True
                                                break
                                    if section_found:
                                        break
                        
                        # If this is the very first procedure and no section set yet
                        if proc_num == 1 and not current_section:
                            # Look for section header above first procedure
                            for check_idx in range(max(0, idx - 5), idx):
                                if check_idx < len(df):
                                    check_row = df.iloc[check_idx]
                                    for cell_val in check_row:
                                        if not pd.isna(cell_val):
                                            cell_str = str(cell_val).strip()
                                            if (len(cell_str) > 3 and len(cell_str) < 50 and 
                                                not cell_str.isdigit() and
                                                not any(word in cell_str.lower() for word in ['inspect', 'replace', 'check', 'test', 'measure', 'clean'])):
                                                current_section = cell_str
                                                if current_section not in [s['name'] for s in self.sections]:
                                                    self.sections.append({
                                                        'name': current_section,
                                                        'row': check_idx,
                                                        'col': 0
                                                    })
                                                break
                                    if current_section:
                                        break
                        
                        self.procedures.append({
                            'section': current_section,
                            'number': proc_num,
                            'description': proc_desc,
                            'full_text': f"{proc_num}. {proc_desc}",
                            'row': idx,
                            'col': 1  # Procedure is in column B
                        })
                        last_proc_num = proc_num
                
                # Also check for procedures in single-cell format (fallback)
                for col_idx, cell_str in enumerate(row_values):
                    if not cell_str:
                        continue
                    
                    procedure_match = re.match(r'^(\d+)\.\s*(.+)', cell_str)
                    if procedure_match:
                        proc_num = int(procedure_match.group(1))
                        proc_desc = procedure_match.group(2)
                        
                        # Avoid duplicates from table format detection
                        if not any(p['number'] == proc_num and p['description'] == proc_desc for p in self.procedures):
                            # Detect section change for single-cell format too
                            if proc_num == 1 and last_proc_num > 1:
                                for check_idx in range(max(0, idx - 3), idx):
                                    if check_idx < len(df):
                                        check_row = df.iloc[check_idx]
                                        for cell_val in check_row:
                                            if not pd.isna(cell_val):
                                                cell_str_check = str(cell_val).strip()
                                                if (len(cell_str_check) > 3 and len(cell_str_check) < 50 and 
                                                    not cell_str_check.isdigit() and
                                                    not re.match(r'^\d+\.', cell_str_check)):
                                                    current_section = cell_str_check
                                                    if current_section not in [s['name'] for s in self.sections]:
                                                        self.sections.append({
                                                            'name': current_section,
                                                            'row': check_idx,
                                                            'col': col_idx
                                                        })
                                                    break
                                        if current_section:
                                            break
                            
                            self.procedures.append({
                                'section': current_section,
                                'number': proc_num,
                                'description': proc_desc,
                                'full_text': cell_str,
                                'row': idx,
                                'col': col_idx
                            })
                            last_proc_num = proc_num
            
            # Display analysis results
            self.analysis_text.delete(1.0, tk.END)
            self.analysis_text.insert(tk.END, f"=== ANALYSIS RESULTS ===\n\n")
            self.analysis_text.insert(tk.END, f"Sheet: {self.selected_sheet}\n")
            self.analysis_text.insert(tk.END, f"Username: {self.user_name}\n\n")
            
            self.analysis_text.insert(tk.END, f"Found {len(self.sections)} sections:\n")
            for section in self.sections:
                self.analysis_text.insert(tk.END, f"  - {section['name']}\n")
            
            self.analysis_text.insert(tk.END, f"\nFound {len(self.procedures)} procedures:\n")
            for proc in self.procedures[:10]:  # Show first 10
                self.analysis_text.insert(tk.END, f"  {proc['number']}. {proc['description'][:50]}...\n")
            
            if len(self.procedures) > 10:
                self.analysis_text.insert(tk.END, f"  ... and {len(self.procedures) - 10} more procedures\n")
            
            if self.procedures:
                self.generate_btn.config(state="normal")
            else:
                self.analysis_text.insert(tk.END, "\nNo procedures found. Please check sheet structure.\n")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to analyze sheet: {str(e)}")
    
    def configure_generation(self):
        """Open configuration dialog for LOV assignment"""
        if not self.procedures:
            return
        
        config_window = ProcedureConfigWindow(self.root, self.procedures, self.lov_database, self.generate_forms, self.base_lov_code)
        config_window.show()
    
    def generate_forms(self, configured_procedures):
        """Generate the 4 Excel files"""
        try:
            self.progress['value'] = 0
            self.root.update()
            
            # Generate base filename
            base_filename = Path(self.source_file).stem
            sheet_clean = re.sub(r'[^\w\-_]', '_', self.selected_sheet)
            
            # Generate FORMHEAD
            formhead_data = self.generate_formhead(base_filename, sheet_clean)
            
            # Generate FORMMENU  
            formmenu_data = self.generate_formmenu(base_filename, sheet_clean)
            
            # Generate FORMTEMPLATE
            formtemplate_data = self.generate_formtemplate(base_filename, sheet_clean, configured_procedures)
            
            # Generate FORMLOV
            formlov_data = self.generate_formlov(configured_procedures)
            
            self.progress['value'] = 50
            self.root.update()
            
            # Save files
            output_dir = Path(self.source_file).parent / f"{base_filename}_{sheet_clean}_Output"
            output_dir.mkdir(exist_ok=True)
            
            # Save each file
            formhead_df = pd.DataFrame([formhead_data])
            formhead_df.to_excel(output_dir / "FORMHEAD.xlsx", index=False)
            
            formmenu_df = pd.DataFrame([formmenu_data])
            formmenu_df.to_excel(output_dir / "FORMMENU.xlsx", index=False)
            
            formtemplate_df = pd.DataFrame(formtemplate_data)
            formtemplate_df.to_excel(output_dir / "FORMTEMPLATE.xlsx", index=False)
            
            formlov_df = pd.DataFrame(formlov_data)
            formlov_df.to_excel(output_dir / "FORMLOV.xlsx", index=False)
            
            self.progress['value'] = 100
            self.root.update()
            
            messagebox.showinfo("Success", f"Forms generated successfully!\nOutput directory: {output_dir}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate forms: {str(e)}")
        finally:
            self.progress['value'] = 0
    
    def generate_formhead(self, base_filename, sheet_clean):
        """Generate FORMHEAD data"""
        form_name = f"{base_filename.upper().replace(' ', '-')}-{sheet_clean.upper()}"
        form_description = self.sections[0]['name'] if self.sections else self.selected_sheet
        
        return {
            'FORMNAME': form_name,
            'VERSION': 1,
            'ENABLE': 1,
            'WFID': 0,
            'FORMDESCRIPTION': form_description,
            'MAPTOPERMITID': '',
            'CATEGORY': 'BASIC',
            'MODIFIEDBY': self.user_name,
            'MODIFIEDDATE': '',
            'STATUS': 'DRAFT',
            'USERNAME': self.user_name,
            'CREATEDATE': '',
            'HEADLINE': '',
            'DETAIL_INFORMATION': ''
        }
    
    def generate_formmenu(self, base_filename, sheet_clean):
        """Generate FORMMENU data"""
        form_name = f"{base_filename.upper().replace(' ', '-')}-{sheet_clean.upper()}"
        form_description = self.sections[0]['name'] if self.sections else self.selected_sheet
        
        return {
            'MNID': '',
            'MNTYPE': 'FORM',
            'MNLABEL': form_description,
            'MNICON': 'ic_survey_general.png',
            'MNDESC': form_description,
            'MNGROUP': '',
            'MNCATEGORY': '',
            'PARENTMNID': 333,
            'FORMNAME': form_name,
            'ATTRIBUTE1': '',
            'ATTRIBUTE2': '',
            'ATTRIBUTE3': '',
            'ISACTIVE': 1,
            'VALIDFROM': '',
            'VALIDTO': ''
        }
    
    def generate_formtemplate(self, base_filename, sheet_clean, configured_procedures):
        """Generate FORMTEMPLATE data"""
        form_name = f"{base_filename.upper().replace(' ', '-')}-{sheet_clean.upper()}"
        form_description = self.sections[0]['name'] if self.sections else self.selected_sheet
        template_data = []
        
        # Generate base identifier
        base_id = f"YK{base_filename[:4].upper()}{sheet_clean[:3].upper()}"
        
        display_option = 0
        current_section = None
        
        # Email field (once at the beginning)
        template_data.append({
            'ORG': 2100,
            'FORMNAME': form_name,
            'KEYNAME': f'{base_id}-TEXSTR0',
            'PARENTKEY': '',
            'KEYTYPE': 'TEXTBOX',
            'KEYDATATYPE': 'STRING',
            'KEYLOV': '',
            'KEYLABEL': 'Email (hanya bisa email pertamina)',
            'KEYFORMULA': 'user.email',
            'KEYHELP': '',
            'KEYHINT': '',
            'DISPLAYOPTION': display_option,
            'VERSION': 1,
            'ENABLE': 1,
            'LASTUPDATEBY': '',
            'LASTUPDATE': '',
            'REQUIRED': 1,
            'SHOWONVALUE': 1,
            'EDITABLE': 1,
            'SHOWONEMPTY': 1,
            'ADDCLASS': '',
            'SHOWONREPORT': 1,
            'CUSTOMLOV': ''
        })
        display_option += 10
        
        # Main header
        template_data.append({
            'ORG': 2100,
            'FORMNAME': form_name,
            'KEYNAME': f'{base_id}-LABSTR0',
            'PARENTKEY': '',
            'KEYTYPE': 'LABEL',
            'KEYDATATYPE': 'STRING',
            'KEYLOV': '',
            'KEYLABEL': f'{form_description} 8000HRS',
            'KEYFORMULA': '',
            'KEYHELP': '',
            'KEYHINT': '',
            'DISPLAYOPTION': display_option,
            'VERSION': 1,
            'ENABLE': 1,
            'LASTUPDATEBY': '',
            'LASTUPDATE': '',
            'REQUIRED': 1,
            'SHOWONVALUE': 1,
            'EDITABLE': 1,
            'SHOWONEMPTY': 1,
            'ADDCLASS': '',
            'SHOWONREPORT': 1,
            'CUSTOMLOV': ''
        })
        display_option += 10
        
        str_counter = 1
        che_counter = 1
        fil_counter = 0
        
        for proc in configured_procedures:
            # Check if new section
            if proc['section'] != current_section and proc['section']:
                current_section = proc['section']
                template_data.append({
                    'ORG': 2100,
                    'FORMNAME': form_name,
                    'KEYNAME': f'{base_id}-LABSTR{str_counter}',
                    'PARENTKEY': '',
                    'KEYTYPE': 'LABEL',
                    'KEYDATATYPE': 'STRING',
                    'KEYLOV': '',
                    'KEYLABEL': proc['section'],
                    'KEYFORMULA': '',
                    'KEYHELP': '',
                    'KEYHINT': '',
                    'DISPLAYOPTION': display_option,
                    'VERSION': 1,
                    'ENABLE': 1,
                    'LASTUPDATEBY': '',
                    'LASTUPDATE': '',
                    'REQUIRED': 1,
                    'SHOWONVALUE': 1,
                    'EDITABLE': 1,
                    'SHOWONEMPTY': 1,
                    'ADDCLASS': '',
                    'SHOWONREPORT': 1,
                    'CUSTOMLOV': ''
                })
                str_counter += 1
                display_option += 10
            
            # Procedure label
            template_data.append({
                'ORG': 2100,
                'FORMNAME': form_name,
                'KEYNAME': f'{base_id}-LABSTR{str_counter}',
                'PARENTKEY': '',
                'KEYTYPE': 'LABEL',
                'KEYDATATYPE': 'STRING',
                'KEYLOV': '',
                'KEYLABEL': proc['full_text'],
                'KEYFORMULA': '',
                'KEYHELP': '',
                'KEYHINT': '',
                'DISPLAYOPTION': display_option,
                'VERSION': 1,
                'ENABLE': 1,
                'LASTUPDATEBY': '',
                'LASTUPDATE': '',
                'REQUIRED': 1,
                'SHOWONVALUE': 1,
                'EDITABLE': 1,
                'SHOWONEMPTY': 1,
                'ADDCLASS': '',
                'SHOWONREPORT': 1,
                'CUSTOMLOV': ''
            })
            str_counter += 1
            display_option += 10
            
            # Standard procedure fields
            fields = [
                ('LISSTR', 'LIST', 'STRING', 'YKN-CPP2-G-603-YN', 'Choose'),
                ('TEXSTR', 'TEXTBOX', 'STRING', '', 'Remarks'),
                ('LISCHE', 'LIST', 'CHECKBOX', proc['condition_lov'], 'Condition found'),
                ('LISCHE', 'LIST', 'CHECKBOX', proc['action_lov'], 'Corrective Action'),
                ('LISSTR', 'LIST', 'STRING', 'YKN-CPP2-G-603-GFB', 'As Left (Good, Fair, Bad)'),
                ('TEXSTR', 'TEXTBOX', 'STRING', '', 'Remarks')
            ]
            
            for field_type, keytype, datatype, keylov, keylabel in fields:
                if field_type == 'LISCHE':
                    counter_suffix = che_counter
                    che_counter += 1
                else:
                    counter_suffix = str_counter
                    str_counter += 1
                
                template_data.append({
                    'ORG': 2100,
                    'FORMNAME': form_name,
                    'KEYNAME': f'{base_id}-{field_type}{counter_suffix}',
                    'PARENTKEY': '',
                    'KEYTYPE': keytype,
                    'KEYDATATYPE': datatype,
                    'KEYLOV': keylov,
                    'KEYLABEL': keylabel,
                    'KEYFORMULA': '',
                    'KEYHELP': '',
                    'KEYHINT': '',
                    'DISPLAYOPTION': display_option,
                    'VERSION': 1,
                    'ENABLE': 1,
                    'LASTUPDATEBY': '',
                    'LASTUPDATE': '',
                    'REQUIRED': 1,
                    'SHOWONVALUE': 1,
                    'EDITABLE': 1,
                    'SHOWONEMPTY': 1,
                    'ADDCLASS': '',
                    'SHOWONREPORT': 1,
                    'CUSTOMLOV': ''
                })
                display_option += 10
            
            # Hidden and File fields
            hidden_key = f'{base_id}-HIDSTR{fil_counter}'
            file_key = f'{base_id}-FILSTR{fil_counter}'
            
            template_data.append({
                'ORG': 2100,
                'FORMNAME': form_name,
                'KEYNAME': hidden_key,
                'PARENTKEY': file_key,
                'KEYTYPE': 'HIDDEN',
                'KEYDATATYPE': 'STRING',
                'KEYLOV': '',
                'KEYLABEL': f'{form_name} UPLOAD FILE',
                'KEYFORMULA': '',
                'KEYHELP': '',
                'KEYHINT': '',
                'DISPLAYOPTION': display_option,
                'VERSION': 1,
                'ENABLE': 1,
                'LASTUPDATEBY': '',
                'LASTUPDATE': '',
                'REQUIRED': 1,
                'SHOWONVALUE': 1,
                'EDITABLE': 1,
                'SHOWONEMPTY': 1,
                'ADDCLASS': '',
                'SHOWONREPORT': 1,
                'CUSTOMLOV': ''
            })
            display_option += 10
            
            template_data.append({
                'ORG': 2100,
                'FORMNAME': form_name,
                'KEYNAME': file_key,
                'PARENTKEY': '',
                'KEYTYPE': 'FILE',
                'KEYDATATYPE': 'STRING',
                'KEYLOV': '',
                'KEYLABEL': 'Silahkan Upload file Pendukung Anda',
                'KEYFORMULA': '',
                'KEYHELP': '',
                'KEYHINT': '',
                'DISPLAYOPTION': display_option,
                'VERSION': 1,
                'ENABLE': 1,
                'LASTUPDATEBY': '',
                'LASTUPDATE': '',
                'REQUIRED': 1,
                'SHOWONVALUE': 1,
                'EDITABLE': 1,
                'SHOWONEMPTY': 1,
                'ADDCLASS': '',
                'SHOWONREPORT': 1,
                'CUSTOMLOV': ''
            })
            display_option += 10
            fil_counter += 1
        
        return template_data
    
    def generate_formlov(self, configured_procedures):
        """Generate FORMLOV data"""
        lov_data = []
        used_lovs = set()
        
        # Add standard LOVs
        standard_lovs = ['YKN-CPP2-G-603-YN', 'YKN-CPP2-G-603-GFB']
        
        for proc in configured_procedures:
            if proc['condition_lov']:
                used_lovs.add(proc['condition_lov'])
            if proc['action_lov']:
                used_lovs.add(proc['action_lov'])
        
        used_lovs.update(standard_lovs)
        
        for lov_code in used_lovs:
            if lov_code in self.lov_database:
                for value in self.lov_database[lov_code]:
                    lov_data.append({
                        'LOVID': '',
                        'ORG': 2100,
                        'LOVNAME': lov_code,
                        'VALUE': value,
                        'VALLOW': '',
                        'VALHI': '',
                        'VALDESC': value,
                        'ENABLE': 1,
                        'TYPE': 'CONFIG'
                    })
        
        return lov_data


class ProcedureConfigWindow:
    def __init__(self, parent, procedures, lov_database, callback):
        self.parent = parent
        self.procedures = procedures
        self.lov_database = lov_database
        self.callback = callback
        self.configured_procedures = []
        
        self.window = None
    
    def show(self):
        """Show configuration window"""
        self.window = tk.Toplevel(self.parent)
        self.window.title("Configure Procedures")
        self.window.geometry("1200x600")
        
        # Make sure window appears on top and gets focus
        self.window.transient(self.parent)
        self.window.grab_set()
        self.window.lift()
        self.window.focus_force()
        
        # Center the window
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (1200 // 2)
        y = (self.window.winfo_screenheight() // 2) - (600 // 2)
        self.window.geometry(f"1200x600+{x}+{y}")
        
        # Create main frame
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Instructions
        instructions = ttk.Label(main_frame, 
                               text="Configure LOV codes for each procedure. Select existing codes or create custom ones.")
        instructions.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Create treeview for procedures
        tree_frame = ttk.Frame(main_frame)
        tree_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        columns = ('Procedure', 'Condition LOV', 'Action LOV')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)
        
        # Configure columns
        self.tree.heading('Procedure', text='Procedure')
        self.tree.heading('Condition LOV', text='Condition Found LOV')
        self.tree.heading('Action LOV', text='Corrective Action LOV')
        
        self.tree.column('Procedure', width=400)
        self.tree.column('Condition LOV', width=200)
        self.tree.column('Action LOV', width=200)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Populate tree with procedures
        for i, proc in enumerate(self.procedures):
            item_id = self.tree.insert('', 'end', values=(
                proc['full_text'],
                'Select...',
                'Select...'
            ))
            # Store procedure data with the item
            self.tree.set(item_id, 'proc_data', i)
        
        # Bind double-click event
        self.tree.bind('<Double-1>', self.on_item_double_click)
        
        # Control buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=2, column=0, pady=(10, 0))
        
        ttk.Button(btn_frame, text="Auto-Configure Similar", 
                  command=self.auto_configure).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="Generate Forms", 
                  command=self.generate_forms).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="Cancel", 
                  command=self.window.destroy).pack(side=tk.LEFT)
        
        # Initialize configured procedures
        for proc in self.procedures:
            self.configured_procedures.append({
                'section': proc['section'],
                'number': proc['number'],
                'description': proc['description'],
                'full_text': proc['full_text'],
                'condition_lov': '',
                'action_lov': ''
            })
        
        # Make sure window is visible and focused
        self.window.deiconify()
        self.window.lift()
        self.window.attributes('-topmost', True)
        self.window.after_idle(lambda: self.window.attributes('-topmost', False))
    
    def on_item_double_click(self, event):
        """Handle double-click on tree item"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        values = self.tree.item(item, 'values')
        proc_index = int(self.tree.item(item, 'values')[0].split('.')[0]) - 1
        
        if proc_index < len(self.configured_procedures):
            self.configure_procedure(item, proc_index)
    
    def configure_procedure(self, tree_item, proc_index):
        """Configure LOV for a specific procedure"""
        proc = self.configured_procedures[proc_index]
        
        config_window = tk.Toplevel(self.window)
        config_window.title(f"Configure: {proc['description'][:50]}...")
        config_window.geometry("500x400")
        config_window.transient(self.window)
        config_window.grab_set()
        
        frame = ttk.Frame(config_window, padding="10")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        config_window.columnconfigure(0, weight=1)
        config_window.rowconfigure(0, weight=1)
        
        # Procedure info
        ttk.Label(frame, text="Procedure:", font=('TkDefaultFont', 10, 'bold')).grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        ttk.Label(frame, text=proc['full_text'], wraplength=400).grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        
        # Condition Found LOV
        ttk.Label(frame, text="Condition Found LOV:").grid(row=2, column=0, sticky=tk.W, pady=(0, 5))
        condition_var = tk.StringVar(value=proc['condition_lov'] if proc['condition_lov'] else 'Select...')
        condition_combo = ttk.Combobox(frame, textvariable=condition_var, width=30)
        condition_combo['values'] = ['Select...'] + list(self.lov_database.keys()) + ['Custom...']
        condition_combo.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Action LOV
        ttk.Label(frame, text="Corrective Action LOV:").grid(row=4, column=0, sticky=tk.W, pady=(0, 5))
        action_var = tk.StringVar(value=proc['action_lov'] if proc['action_lov'] else 'Select...')
        action_combo = ttk.Combobox(frame, textvariable=action_var, width=30)
        action_combo['values'] = ['Select...'] + list(self.lov_database.keys()) + ['Custom...']
        action_combo.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        
        frame.columnconfigure(1, weight=1)
        
        def save_config():
            condition_lov = condition_var.get() if condition_var.get() != 'Select...' else ''
            action_lov = action_var.get() if action_var.get() != 'Select...' else ''
            
            # Handle custom LOV
            if condition_lov == 'Custom...':
                condition_lov = simpledialog.askstring("Custom LOV", "Enter custom condition LOV code:")
                if condition_lov:
                    # Add to database for this session
                    values = simpledialog.askstring("LOV Values", "Enter comma-separated values:").split(',')
                    self.lov_database[condition_lov] = [v.strip() for v in values]
            
            if action_lov == 'Custom...':
                action_lov = simpledialog.askstring("Custom LOV", "Enter custom action LOV code:")
                if action_lov:
                    values = simpledialog.askstring("LOV Values", "Enter comma-separated values:").split(',')
                    self.lov_database[action_lov] = [v.strip() for v in values]
            
            # Update configured procedure
            self.configured_procedures[proc_index]['condition_lov'] = condition_lov
            self.configured_procedures[proc_index]['action_lov'] = action_lov
            
            # Update tree display
            self.tree.item(tree_item, values=(
                proc['full_text'],
                condition_lov if condition_lov else 'Select...',
                action_lov if action_lov else 'Select...'
            ))
            
            config_window.destroy()
        
        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=6, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Button(btn_frame, text="Save", command=save_config).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="Cancel", command=config_window.destroy).pack(side=tk.LEFT)
    
    def auto_configure(self):
        """Auto-configure similar procedures"""
        # Simple auto-configuration based on keywords
        for proc in self.configured_procedures:
            desc_lower = proc['description'].lower()
            
            # Default condition LOV based on keywords
            if any(word in desc_lower for word in ['filter', 'air']):
                proc['condition_lov'] = 'YKN-CPP2-G-603-DCG'  # Dirty, Clogged, Good
            elif any(word in desc_lower for word in ['oil', 'lube']):
                proc['condition_lov'] = 'YKN-CPP2-G-603-DLG'  # Dirty, Low Level, Good
            elif any(word in desc_lower for word in ['check', 'visual']):
                proc['condition_lov'] = 'YKN-CPP2-G-603-LG'   # Leak, Good
            else:
                proc['condition_lov'] = 'YKN-CPP2-G-603-CNBG' # Clogged, Not Working, Broken, Good
            
            # Default action LOV
            if 'replace' in desc_lower:
                proc['action_lov'] = 'YKN-CPP2-G-603-R1'     # Replaced
            elif any(word in desc_lower for word in ['check', 'visual']):
                proc['action_lov'] = 'YKN-CPP2-G-603-RR2'    # Repaired, Replaced
            else:
                proc['action_lov'] = 'YKN-CPP2-G-603-RR1'    # Reset, Replaced
        
        # Update tree display
        for i, item in enumerate(self.tree.get_children()):
            proc = self.configured_procedures[i]
            self.tree.item(item, values=(
                proc['full_text'],
                proc['condition_lov'],
                proc['action_lov']
            ))
        
        messagebox.showinfo("Auto-Configure", "Auto-configuration completed. Please review and adjust as needed.")
    
    def generate_forms(self):
        """Generate forms with configured procedures"""
        # Validate that all procedures have LOV codes
        missing_configs = []
        for i, proc in enumerate(self.configured_procedures):
            if not proc['condition_lov'] or not proc['action_lov']:
                missing_configs.append(f"{proc['number']}. {proc['description'][:30]}...")
        
        if missing_configs:
            messagebox.showwarning("Incomplete Configuration", 
                                 f"Please configure LOV codes for all procedures:\n" + 
                                 "\n".join(missing_configs[:5]) + 
                                 (f"\n... and {len(missing_configs) - 5} more" if len(missing_configs) > 5 else ""))
            return
        
        self.window.destroy()
        self.callback(self.configured_procedures)


def main():
    root = tk.Tk()
    app = FormGeneratorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()