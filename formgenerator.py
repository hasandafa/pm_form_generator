import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
import os
import re
from datetime import datetime
from pathlib import Path
import json
import hashlib
from collections import defaultdict

class ExcelFormatHandler:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Format Handler v3.0 - Focused Solution")
        self.root.geometry("1200x700")
        
        # Core variables
        self.source_file = None
        self.selected_sheet = None
        self.raw_dataframe = None
        self.parsed_procedures = []
        self.user_name = ""
        
        # Unique identifier system
        self.unique_tracker_file = "unique_identifiers.json"
        self.used_identifiers = self.load_identifier_tracker()
        self.current_prefixes = {}
        
        # LOV configuration storage
        self.procedure_lov_config = {}
        
        # Format detection patterns (simplified)
        self.format_keywords = {
            'maintenance': ['procedure', 'condition', 'corrective', 'inspect', 'replace', 'clean'],
            'checklist': ['check', 'ok', 'not ok', 'startup', 'parameter'],
            'calibration': ['calibrate', 'tolerance', 'as found', 'as left', 'standard'],
            'monitoring': ['monitor', 'pressure', 'temperature', 'before service', 'after service']
        }
        
        self.create_interface()
    
    def load_identifier_tracker(self):
        """Load unique identifier tracking database"""
        try:
            if os.path.exists(self.unique_tracker_file):
                with open(self.unique_tracker_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return {
                "prefixes": [],
                "lov_codes": [],
                "form_names": [],
                "template_ids": []
            }
        except Exception as e:
            print(f"Error loading identifier tracker: {e}")
            return {"prefixes": [], "lov_codes": [], "form_names": [], "template_ids": []}
    
    def save_identifier_tracker(self):
        """Save unique identifier tracking database"""
        try:
            with open(self.unique_tracker_file, 'w', encoding='utf-8') as f:
                json.dump(self.used_identifiers, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving identifier tracker: {e}")
    
    def generate_unique_prefix(self, base_text, prefix_type="general"):
        """Generate unique prefix with conflict resolution"""
        # Clean and truncate base text
        clean_text = re.sub(r'[^\w]', '', base_text.upper())[:8]
        
        # Create base prefix (first 3 chars + hash)
        if len(clean_text) >= 3:
            base_prefix = clean_text[:3]
        else:
            base_prefix = (clean_text + "XXX")[:3]
        
        # Add hash for uniqueness
        hash_part = hashlib.md5(base_text.encode()).hexdigest()[:2].upper()
        candidate = f"{base_prefix}-{hash_part}"
        
        # Check for conflicts and resolve
        used_list = self.used_identifiers.get("prefixes", [])
        counter = 1
        original_candidate = candidate
        
        while candidate in used_list:
            candidate = f"{original_candidate}-{counter:02d}"
            counter += 1
            if counter > 99:  # Fallback to timestamp
                candidate = f"{base_prefix}-{datetime.now().strftime('%H%M')}"
                break
        
        # Record the new prefix
        used_list.append(candidate)
        self.used_identifiers["prefixes"] = used_list
        
        return candidate
    
    def generate_lov_code(self, values_text, procedure_index=None):
        """Generate LOV code based on values with proper uniqueness"""
        if not values_text:
            return f"DEFAULT-{procedure_index or '00'}"
        
        # Parse values and create code
        values = [v.strip().upper() for v in values_text.split(',') if v.strip()]
        
        # Create signature from first letters
        code_parts = []
        for value in values[:3]:  # Use max 3 values
            if value:
                code_parts.append(value[0])
        
        base_code = ''.join(code_parts) if code_parts else "DEF"
        
        # Add current sheet prefix if available
        if hasattr(self, 'sheet_prefix') and self.sheet_prefix:
            full_code = f"{self.sheet_prefix}-{base_code}"
        else:
            full_code = base_code
        
        # Ensure uniqueness
        used_lov_codes = self.used_identifiers.get("lov_codes", [])
        counter = 1
        original_code = full_code
        
        while full_code in used_lov_codes:
            full_code = f"{original_code}{counter}"
            counter += 1
        
        # Record the new LOV code
        used_lov_codes.append(full_code)
        self.used_identifiers["lov_codes"] = used_lov_codes
        
        return full_code
    
    def create_interface(self):
        """Create clean, focused interface"""
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Step 1: File Selection
        self.create_file_section(main_frame)
        
        # Step 2: Analysis Results
        self.create_analysis_section(main_frame)
        
        # Step 3: LOV Configuration
        self.create_lov_section(main_frame)
        
        # Step 4: Generate Output
        self.create_output_section(main_frame)
        
        # Status bar
        self.status_bar = ttk.Label(self.root, text="Ready - Select Excel file to begin", relief=tk.SUNKEN)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def create_file_section(self, parent):
        """Create file selection section"""
        file_frame = ttk.LabelFrame(parent, text="Step 1: File Selection", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        # File selection row
        file_row = ttk.Frame(file_frame)
        file_row.pack(fill=tk.X, pady=5)
        
        ttk.Button(file_row, text="Select Excel File", command=self.select_file).pack(side=tk.LEFT)
        self.file_label = ttk.Label(file_row, text="No file selected", foreground="gray")
        self.file_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # User and sheet row
        config_row = ttk.Frame(file_frame)
        config_row.pack(fill=tk.X, pady=5)
        
        ttk.Label(config_row, text="User:").pack(side=tk.LEFT)
        
        # User selection frame
        user_frame = ttk.Frame(config_row)
        user_frame.pack(side=tk.LEFT, padx=(5, 20))
        
        # Default user option
        self.user_mode = tk.StringVar(value="default")
        default_radio = ttk.Radiobutton(user_frame, text="MK.ABDULLAH.DAFA", 
                                       variable=self.user_mode, value="default")
        default_radio.pack(side=tk.LEFT)
        
        # Custom user option
        custom_radio = ttk.Radiobutton(user_frame, text="Custom:", 
                                      variable=self.user_mode, value="custom")
        custom_radio.pack(side=tk.LEFT, padx=(10, 5))
        
        # Custom user entry
        self.user_entry = ttk.Entry(user_frame, width=15, state="disabled")
        self.user_entry.pack(side=tk.LEFT, padx=(5, 0))
        
        # Bind radio button changes to enable/disable entry
        self.user_mode.trace('w', self.on_user_mode_change)
        
        ttk.Label(config_row, text="Sheet:").pack(side=tk.LEFT, padx=(20, 0))
        self.sheet_combo = ttk.Combobox(config_row, width=25, state="readonly")
        self.sheet_combo.pack(side=tk.LEFT, padx=(5, 20))
        
        ttk.Button(config_row, text="Analyze Format", command=self.analyze_format).pack(side=tk.LEFT, padx=(10, 0))
        
        # Unique identifiers display
        self.identifiers_label = ttk.Label(file_frame, text="Generated Prefixes: None", 
                                         font=('Consolas', 9), foreground="blue")
        self.identifiers_label.pack(anchor=tk.W, pady=(5, 0))
    
    def on_user_mode_change(self, *args):
        """Handle user mode radio button changes"""
        if self.user_mode.get() == "default":
            self.user_entry.config(state="disabled")
            self.user_entry.delete(0, tk.END)
        else:
            self.user_entry.config(state="normal")
            self.user_entry.focus()
    
    def get_current_user(self):
        """Get the current user based on selected mode"""
        if self.user_mode.get() == "default":
            return "MK.ABDULLAH.DAFA"
        else:
            return self.user_entry.get().strip() or "Unknown User"
    
    def create_analysis_section(self, parent):
        """Create analysis results section"""
        analysis_frame = ttk.LabelFrame(parent, text="Step 2: Analysis Results", padding="10")
        analysis_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Results display
        self.results_text = ScrolledText(analysis_frame, height=12, font=('Consolas', 9))
        self.results_text.pack(fill=tk.BOTH, expand=True)
        
        # Analysis controls
        controls_frame = ttk.Frame(analysis_frame)
        controls_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(controls_frame, text="Accept & Configure LOVs", 
                  command=self.accept_and_configure).pack(side=tk.LEFT)
        ttk.Button(controls_frame, text="Re-analyze", 
                  command=self.analyze_format).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Button(controls_frame, text="View Raw Data", 
                  command=self.show_raw_data).pack(side=tk.LEFT, padx=(10, 0))
    
    def create_lov_section(self, parent):
        """Create LOV configuration section"""
        lov_frame = ttk.LabelFrame(parent, text="Step 3: LOV Configuration", padding="10")
        lov_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Instructions
        ttk.Label(lov_frame, text="Configure condition and action values for procedures (comma-separated):",
                 font=('TkDefaultFont', 9)).pack(anchor=tk.W)
        
        # LOV configuration area (initially hidden)
        self.lov_config_frame = ttk.Frame(lov_frame)
        self.lov_config_frame.pack(fill=tk.X, pady=10)
        
        # This will be populated when procedures are analyzed
        self.lov_widgets = []
        
        # LOV controls
        lov_controls = ttk.Frame(lov_frame)
        lov_controls.pack(fill=tk.X)
        
        ttk.Button(lov_controls, text="Auto-Configure Common LOVs", 
                  command=self.auto_configure_lovs).pack(side=tk.LEFT)
        ttk.Button(lov_controls, text="Clear All LOVs", 
                  command=self.clear_lov_config).pack(side=tk.LEFT, padx=(10, 0))
    
    def create_output_section(self, parent):
        """Create output generation section"""
        output_frame = ttk.LabelFrame(parent, text="Step 4: Generate Forms", padding="10")
        output_frame.pack(fill=tk.X)
        
        # Output directory
        dir_frame = ttk.Frame(output_frame)
        dir_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(dir_frame, text="Output Directory:").pack(side=tk.LEFT)
        self.output_dir = tk.StringVar(value=os.getcwd())
        ttk.Entry(dir_frame, textvariable=self.output_dir, width=50).pack(side=tk.LEFT, padx=(10, 10))
        ttk.Button(dir_frame, text="Browse", command=self.select_output_dir).pack(side=tk.LEFT)
        
        # Generation controls
        gen_frame = ttk.Frame(output_frame)
        gen_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(gen_frame, text="Generate Excel Forms", 
                  command=self.generate_forms, style='Accent.TButton').pack(side=tk.LEFT)
        ttk.Button(gen_frame, text="Preview Generation", 
                  command=self.preview_generation).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Button(gen_frame, text="Save Configuration", 
                  command=self.save_config).pack(side=tk.LEFT, padx=(10, 0))
    
    def select_file(self):
        """Select Excel file and generate file prefix"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            self.source_file = file_path
            filename = Path(file_path).stem
            
            # Generate unique file prefix
            self.file_prefix = self.generate_unique_prefix(filename, "file")
            self.current_prefixes['file'] = self.file_prefix
            
            self.file_label.config(text=os.path.basename(file_path), foreground="black")
            self.load_sheets()
            self.update_identifiers_display()
            self.status_bar.config(text=f"File loaded: {os.path.basename(file_path)}")
    
    def load_sheets(self):
        """Load sheets from selected Excel file"""
        try:
            excel_file = pd.ExcelFile(self.source_file)
            self.sheet_combo['values'] = excel_file.sheet_names
            
            if excel_file.sheet_names:
                self.sheet_combo.set(excel_file.sheet_names[0])
                # Generate sheet prefix for first sheet
                self.sheet_prefix = self.generate_unique_prefix(excel_file.sheet_names[0], "sheet")
                self.current_prefixes['sheet'] = self.sheet_prefix
                self.update_identifiers_display()
            
            self.results_text.delete(1.0, tk.END)
            self.results_text.insert(tk.END, f"✅ File loaded successfully\n")
            self.results_text.insert(tk.END, f"📊 Found {len(excel_file.sheet_names)} sheets:\n\n")
            
            for i, sheet in enumerate(excel_file.sheet_names, 1):
                self.results_text.insert(tk.END, f"   {i}. {sheet}\n")
            
            self.results_text.insert(tk.END, f"\n🎯 Click 'Analyze Format' to detect procedures\n")
            
        except Exception as e:
            messagebox.showerror("File Error", f"Cannot read Excel file: {str(e)}")
            self.status_bar.config(text="Error loading file")
    
    def update_identifiers_display(self):
        """Update the identifiers display"""
        prefixes = []
        if hasattr(self, 'file_prefix'):
            prefixes.append(f"File: {self.file_prefix}")
        if hasattr(self, 'sheet_prefix'):
            prefixes.append(f"Sheet: {self.sheet_prefix}")
        
        if prefixes:
            self.identifiers_label.config(text=f"Generated Prefixes: {' | '.join(prefixes)}")
        else:
            self.identifiers_label.config(text="Generated Prefixes: None")
    
    def analyze_format(self):
        """Analyze selected sheet and detect procedures"""
        if not self.source_file or not self.sheet_combo.get():
            messagebox.showwarning("Selection Required", "Please select file and sheet first")
            return
        
        self.selected_sheet = self.sheet_combo.get()
        
        # Update sheet prefix if changed
        if not hasattr(self, 'sheet_prefix') or self.sheet_combo.get() != self.current_prefixes.get('sheet_name'):
            self.sheet_prefix = self.generate_unique_prefix(self.selected_sheet, "sheet")
            self.current_prefixes['sheet'] = self.sheet_prefix
            self.current_prefixes['sheet_name'] = self.selected_sheet
            self.update_identifiers_display()
        
        try:
            self.status_bar.config(text="Analyzing format...")
            
            # Read sheet data
            self.raw_dataframe = pd.read_excel(self.source_file, sheet_name=self.selected_sheet, header=None)
            
            # Detect format and extract procedures
            detected_format = self.detect_format(self.raw_dataframe)
            self.parsed_procedures = self.extract_procedures(self.raw_dataframe, detected_format)
            
            # Display results
            self.display_analysis_results(detected_format)
            
            self.status_bar.config(text=f"Analysis complete - Found {len(self.parsed_procedures)} procedures")
            
        except Exception as e:
            messagebox.showerror("Analysis Error", f"Failed to analyze format: {str(e)}")
            self.status_bar.config(text="Analysis failed")
    
    def detect_format(self, df):
        """Simple but effective format detection"""
        all_text = ""
        for idx, row in df.iterrows():
            for cell in row:
                if not pd.isna(cell):
                    all_text += str(cell).lower() + " "
        
        # Score each format type
        format_scores = {}
        for format_type, keywords in self.format_keywords.items():
            score = sum(1 for keyword in keywords if keyword in all_text)
            format_scores[format_type] = score / len(keywords)  # Normalize
        
        # Return the best match
        if format_scores:
            best_format = max(format_scores, key=format_scores.get)
            confidence = format_scores[best_format]
            return {'type': best_format, 'confidence': confidence, 'scores': format_scores}
        
        return {'type': 'unknown', 'confidence': 0.0, 'scores': {}}
    
    def extract_procedures(self, df, format_info):
        """Extract procedures from dataframe based on detected format"""
        procedures = []
        
        try:
            for idx, row in df.iterrows():
                row_values = [str(cell).strip() if not pd.isna(cell) else '' for cell in row]
                
                # Look for numbered procedures
                for col_idx, cell_text in enumerate(row_values):
                    if self.is_procedure_text(cell_text):
                        procedures.append({
                            'index': len(procedures) + 1,
                            'text': cell_text,
                            'row': idx,
                            'col': col_idx,
                            'format_type': format_info['type']
                        })
                        break  # Only take first procedure per row
        
        except Exception as e:
            print(f"Error extracting procedures: {e}")
        
        return procedures
    
    def is_procedure_text(self, text):
        """Determine if text represents a procedure"""
        if not text or len(text.strip()) < 10:
            return False
        
        text = text.strip()
        
        # Check for numbered procedures
        if re.match(r'^\d+[\.\)]\s*.{10,}', text):
            return True
        
        # Check for procedure indicators
        procedure_indicators = ['inspect', 'check', 'clean', 'replace', 'calibrate', 'test', 'monitor', 'verify']
        text_lower = text.lower()
        
        if any(indicator in text_lower for indicator in procedure_indicators):
            return len(text) > 15  # Must be substantial text
        
        return False
    
    def display_analysis_results(self, format_info):
        """Display analysis results in a clear, actionable format"""
        self.results_text.delete(1.0, tk.END)
        
        self.results_text.insert(tk.END, "📋 FORMAT ANALYSIS RESULTS\n")
        self.results_text.insert(tk.END, "=" * 50 + "\n\n")
        
        # Format detection
        self.results_text.insert(tk.END, f"📁 File: {os.path.basename(self.source_file)}\n")
        self.results_text.insert(tk.END, f"📄 Sheet: {self.selected_sheet}\n")
        self.results_text.insert(tk.END, f"🔍 Detected Format: {format_info['type'].upper()}\n")
        self.results_text.insert(tk.END, f"🎯 Confidence: {format_info['confidence']:.1%}\n\n")
        
        # Format scores breakdown
        if format_info['scores']:
            self.results_text.insert(tk.END, "📊 Format Score Breakdown:\n")
            for fmt, score in format_info['scores'].items():
                self.results_text.insert(tk.END, f"   • {fmt.title()}: {score:.1%}\n")
            self.results_text.insert(tk.END, "\n")
        
        # Procedures found
        if self.parsed_procedures:
            self.results_text.insert(tk.END, f"✅ Found {len(self.parsed_procedures)} procedures:\n")
            self.results_text.insert(tk.END, "-" * 40 + "\n")
            
            for i, proc in enumerate(self.parsed_procedures[:8], 1):  # Show first 8
                procedure_text = proc['text']
                if len(procedure_text) > 60:
                    procedure_text = procedure_text[:60] + "..."
                self.results_text.insert(tk.END, f"{i:2d}. {procedure_text}\n")
            
            if len(self.parsed_procedures) > 8:
                remaining = len(self.parsed_procedures) - 8
                self.results_text.insert(tk.END, f"... and {remaining} more procedures\n")
            
            self.results_text.insert(tk.END, f"\n🎯 Ready for LOV configuration!\n")
            self.results_text.insert(tk.END, "Click 'Accept & Configure LOVs' to proceed.\n")
        else:
            self.results_text.insert(tk.END, "❌ No procedures detected.\n")
            self.results_text.insert(tk.END, "Try 'View Raw Data' to manually identify procedures.\n")
    
    def accept_and_configure(self):
        """Accept analysis results and setup LOV configuration"""
        if not self.parsed_procedures:
            messagebox.showwarning("No Procedures", "No procedures found. Please analyze the sheet first.")
            return
        
        self.setup_lov_configuration()
        messagebox.showinfo("Ready for Configuration", 
                          f"Found {len(self.parsed_procedures)} procedures.\n"
                          "Please configure condition and action values below.")
    
    def setup_lov_configuration(self):
        """Setup LOV configuration interface for detected procedures"""
        # Clear existing widgets
        for widget in self.lov_widgets:
            widget.destroy()
        self.lov_widgets.clear()
        
        # Create scrollable frame for procedures
        canvas = tk.Canvas(self.lov_config_frame, height=200)
        scrollbar = ttk.Scrollbar(self.lov_config_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Headers
        header_frame = ttk.Frame(scrollable_frame)
        header_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(header_frame, text="Procedure", width=40, font=('TkDefaultFont', 9, 'bold')).pack(side=tk.LEFT)
        ttk.Label(header_frame, text="Condition Values", width=20, font=('TkDefaultFont', 9, 'bold')).pack(side=tk.LEFT, padx=5)
        ttk.Label(header_frame, text="Action Values", width=20, font=('TkDefaultFont', 9, 'bold')).pack(side=tk.LEFT, padx=5)
        ttk.Label(header_frame, text="Generated LOV Code", width=25, font=('TkDefaultFont', 9, 'bold')).pack(side=tk.LEFT, padx=5)
        
        # Create procedure configuration rows
        for i, procedure in enumerate(self.parsed_procedures):
            proc_frame = ttk.Frame(scrollable_frame)
            proc_frame.pack(fill=tk.X, padx=5, pady=2)
            
            # Procedure description (truncated)
            proc_text = procedure['text']
            if len(proc_text) > 40:
                proc_text = proc_text[:37] + "..."
            
            proc_label = ttk.Label(proc_frame, text=proc_text, width=40)
            proc_label.pack(side=tk.LEFT)
            
            # Condition values entry
            condition_var = tk.StringVar()
            condition_entry = ttk.Entry(proc_frame, textvariable=condition_var, width=20)
            condition_entry.pack(side=tk.LEFT, padx=5)
            
            # Action values entry
            action_var = tk.StringVar()
            action_entry = ttk.Entry(proc_frame, textvariable=action_var, width=20)
            action_entry.pack(side=tk.LEFT, padx=5)
            
            # Generated LOV code display
            lov_code_var = tk.StringVar(value="Not Generated")
            lov_label = ttk.Label(proc_frame, textvariable=lov_code_var, width=25, foreground="blue")
            lov_label.pack(side=tk.LEFT, padx=5)
            
            # Store references for later use
            proc_config = {
                'procedure': procedure,
                'condition_var': condition_var,
                'action_var': action_var,
                'lov_code_var': lov_code_var,
                'frame': proc_frame
            }
            
            self.procedure_lov_config[i] = proc_config
            
            # Bind events to auto-generate LOV codes
            condition_var.trace('w', lambda name, index, mode, idx=i: self.update_lov_code(idx))
            action_var.trace('w', lambda name, index, mode, idx=i: self.update_lov_code(idx))
            
            self.lov_widgets.extend([proc_frame, proc_label, condition_entry, action_entry, lov_label])
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.lov_widgets.extend([canvas, scrollbar, scrollable_frame])
    
    def update_lov_code(self, procedure_index):
        """Update LOV code when values change"""
        if procedure_index not in self.procedure_lov_config:
            return
        
        config = self.procedure_lov_config[procedure_index]
        condition_values = config['condition_var'].get().strip()
        action_values = config['action_var'].get().strip()
        
        if condition_values or action_values:
            # Generate LOV codes
            if condition_values:
                condition_lov = self.generate_lov_code(condition_values, f"C{procedure_index+1:02d}")
            else:
                condition_lov = ""
            
            if action_values:
                action_lov = self.generate_lov_code(action_values, f"A{procedure_index+1:02d}")
            else:
                action_lov = ""
            
            # Update display
            if condition_lov and action_lov:
                display_text = f"C:{condition_lov} | A:{action_lov}"
            elif condition_lov:
                display_text = f"C:{condition_lov}"
            elif action_lov:
                display_text = f"A:{action_lov}"
            else:
                display_text = "Enter values above"
            
            config['lov_code_var'].set(display_text)
            
            # Store the generated codes
            config['condition_lov'] = condition_lov
            config['action_lov'] = action_lov
        else:
            config['lov_code_var'].set("Enter values above")
    
    def auto_configure_lovs(self):
        """Auto-configure common LOV values based on procedure text"""
        if not self.procedure_lov_config:
            messagebox.showwarning("No Procedures", "Please analyze procedures first")
            return
        
        # Common mappings based on keywords
        condition_mappings = {
            'inspect': 'Good,Dirty,Damaged,Missing',
            'check': 'OK,Not OK,Needs Attention',
            'clean': 'Clean,Dirty,Blocked',
            'replace': 'Good,Worn,Damaged',
            'calibrate': 'In Tolerance,Out of Tolerance',
            'test': 'Pass,Fail',
            'monitor': 'Normal,High,Low'
        }
        
        action_mappings = {
            'inspect': 'No Action,Clean,Repair,Replace',
            'check': 'No Action,Adjust,Repair',
            'clean': 'Cleaned,Replaced',
            'replace': 'Replaced,Repaired',
            'calibrate': 'Calibrated,Adjusted',
            'test': 'No Action,Repaired',
            'monitor': 'No Action,Adjusted'
        }
        
        configured_count = 0
        
        for idx, config in self.procedure_lov_config.items():
            procedure_text = config['procedure']['text'].lower()
            
            # Find matching keyword
            for keyword in condition_mappings:
                if keyword in procedure_text:
                    config['condition_var'].set(condition_mappings[keyword])
                    config['action_var'].set(action_mappings[keyword])
                    configured_count += 1
                    break
            else:
                # Default values if no keyword match
                config['condition_var'].set('Good,Damaged')
                config['action_var'].set('No Action,Repaired')
                configured_count += 1
        
        messagebox.showinfo("Auto-Configuration Complete", 
                          f"Configured LOVs for {configured_count} procedures.\n"
                          "Review and modify as needed before generating forms.")
    
    def clear_lov_config(self):
        """Clear all LOV configurations"""
        for config in self.procedure_lov_config.values():
            config['condition_var'].set('')
            config['action_var'].set('')
    
    def show_raw_data(self):
        """Show raw data structure for manual analysis"""
        if self.raw_dataframe is None:
            messagebox.showwarning("No Data", "Please load and analyze a sheet first")
            return
        
        # Create raw data viewer window
        viewer = tk.Toplevel(self.root)
        viewer.title("Raw Data Structure")
        viewer.geometry("800x600")
        
        text_widget = ScrolledText(viewer, wrap=tk.WORD, font=('Consolas', 9))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        text_widget.insert(tk.END, "RAW DATA STRUCTURE\n")
        text_widget.insert(tk.END, "=" * 60 + "\n\n")
        
        # Show first 25 rows with coordinates
        for idx, row in self.raw_dataframe.head(25).iterrows():
            text_widget.insert(tk.END, f"Row {idx:2d}: ")
            for col_idx, cell in enumerate(row):
                if not pd.isna(cell):
                    cell_str = str(cell).strip()
                    if cell_str:
                        text_widget.insert(tk.END, f"[{col_idx}] {cell_str[:35]:<35} ")
            text_widget.insert(tk.END, "\n")
        
        text_widget.configure(state='disabled')
    
    def select_output_dir(self):
        """Select output directory"""
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_dir.set(directory)
    
    def preview_generation(self):
        """Preview what will be generated"""
        if not self.procedure_lov_config:
            messagebox.showwarning("No Configuration", "Please configure procedures first")
            return
        
        # Create preview window
        preview = tk.Toplevel(self.root)
        preview.title("Generation Preview")
        preview.geometry("700x500")
        
        text_widget = ScrolledText(preview, wrap=tk.WORD, font=('Consolas', 9))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Generate preview content
        preview_content = self.generate_preview_content()
        text_widget.insert(tk.END, preview_content)
        text_widget.configure(state='disabled')
    
    def generate_preview_content(self):
        """Generate preview content for form generation"""
        content = "FORM GENERATION PREVIEW\n"
        content += "=" * 50 + "\n\n"
        
        # File information
        content += f"Source File: {os.path.basename(self.source_file)}\n"
        content += f"Sheet: {self.selected_sheet}\n"
        content += f"User: {self.get_current_user()}\n"
        content += f"File Prefix: {getattr(self, 'file_prefix', 'Not generated')}\n"
        content += f"Sheet Prefix: {getattr(self, 'sheet_prefix', 'Not generated')}\n\n"
        
        # Files to be generated
        content += "FILES TO BE GENERATED:\n"
        content += "✓ FORMHEAD.xlsx - Form metadata and configuration\n"
        content += f"✓ FORMTEMPLATE.xlsx - {len(self.parsed_procedures) * 9} template entries\n"
        content += f"✓ FORMLOV.xlsx - LOV definitions\n"
        content += "✓ FORMMENU.xlsx - Menu structure\n\n"
        
        # Procedure summary
        content += f"PROCEDURES SUMMARY ({len(self.parsed_procedures)}):\n"
        content += "-" * 40 + "\n"
        
        configured_count = 0
        for i, config in self.procedure_lov_config.items():
            condition_vals = config['condition_var'].get()
            action_vals = config['action_var'].get()
            
            if condition_vals or action_vals:
                configured_count += 1
            
            proc_text = config['procedure']['text'][:45]
            status = "✓ Configured" if (condition_vals or action_vals) else "⚠ Not configured"
            
            content += f"{i+1:2d}. {proc_text}... [{status}]\n"
        
        content += f"\nCONFIGURATION STATUS:\n"
        content += f"✓ Configured: {configured_count}/{len(self.procedure_lov_config)}\n"
        
        if configured_count == 0:
            content += "\n⚠ WARNING: No procedures have been configured with LOV values.\n"
            content += "Please configure LOVs before generating forms.\n"
        elif configured_count < len(self.procedure_lov_config):
            content += f"\n⚠ NOTICE: {len(self.procedure_lov_config) - configured_count} procedures not configured.\n"
            content += "These will use default values.\n"
        
        return content
    
    def generate_forms(self):
        """Generate the Excel form files"""
        if not self.procedure_lov_config:
            messagebox.showwarning("No Configuration", "Please analyze and configure procedures first")
            return
        
        # Validate that user has configured at least some procedures
        configured_count = sum(1 for config in self.procedure_lov_config.values() 
                             if config['condition_var'].get() or config['action_var'].get())
        
        if configured_count == 0:
            result = messagebox.askyesno("No LOV Configuration", 
                                       "No procedures have been configured with LOV values.\n"
                                       "Do you want to auto-configure all procedures first?")
            if result:
                self.auto_configure_lovs()
            else:
                return
        
        try:
            output_dir = self.output_dir.get()
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            user_name = self.get_current_user()
            
            # Save identifier tracking
            self.save_identifier_tracker()
            
            # Generate files
            files_created = []
            
            # 1. FORMHEAD.xlsx
            formhead_file = os.path.join(output_dir, f"FORMHEAD_{self.sheet_prefix}_{timestamp}.xlsx")
            self.create_formhead(formhead_file, user_name)
            files_created.append(formhead_file)
            
            # 2. FORMTEMPLATE.xlsx
            template_file = os.path.join(output_dir, f"FORMTEMPLATE_{self.sheet_prefix}_{timestamp}.xlsx")
            self.create_formtemplate(template_file)
            files_created.append(template_file)
            
            # 3. FORMLOV.xlsx
            lov_file = os.path.join(output_dir, f"FORMLOV_{self.sheet_prefix}_{timestamp}.xlsx")
            self.create_formlov(lov_file)
            files_created.append(lov_file)
            
            # 4. FORMMENU.xlsx
            menu_file = os.path.join(output_dir, f"FORMMENU_{self.sheet_prefix}_{timestamp}.xlsx")
            self.create_formmenu(menu_file)
            files_created.append(menu_file)
            
            # Show success message
            success_msg = f"Successfully generated {len(files_created)} files:\n\n"
            success_msg += "\n".join([os.path.basename(f) for f in files_created])
            success_msg += f"\n\nOutput directory: {output_dir}"
            
            messagebox.showinfo("Generation Complete", success_msg)
            
            # Update status
            self.status_bar.config(text=f"Generated {len(files_created)} Excel files successfully")
            
        except Exception as e:
            messagebox.showerror("Generation Error", f"Failed to generate files: {str(e)}")
    
    def create_formhead(self, filename, user_name):
        """Create FORMHEAD.xlsx"""
        data = {
            'FORMNAME': [f"{self.sheet_prefix}-FORM"],
            'TITLE': [f"Maintenance Form - {self.selected_sheet}"],
            'DESCRIPTION': [f"Generated from {os.path.basename(self.source_file)}"],
            'CREATED_BY': [user_name],
            'CREATED_DATE': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
            'SOURCE_FILE': [os.path.basename(self.source_file)],
            'SOURCE_SHEET': [self.selected_sheet],
            'PROCEDURES_COUNT': [len(self.parsed_procedures)],
            'FILE_PREFIX': [self.file_prefix],
            'SHEET_PREFIX': [self.sheet_prefix]
        }
        
        df = pd.DataFrame(data)
        df.to_excel(filename, index=False)
    
    def create_formtemplate(self, filename):
        """Create FORMTEMPLATE.xlsx with 9 entries per procedure"""
        template_data = []
        
        for i, config in self.procedure_lov_config.items():
            procedure = config['procedure']
            base_id = f"{self.sheet_prefix}-P{i+1:03d}"
            
            # 9 template entries per procedure
            entries = [
                {'TEMPLATEID': f"{base_id}-LBL", 'TYPE': 'LABEL', 'DESCRIPTION': procedure['text']},
                {'TEMPLATEID': f"{base_id}-LST", 'TYPE': 'LIST', 'DESCRIPTION': 'Condition Found', 
                 'LOVCODE': config.get('condition_lov', '')},
                {'TEMPLATEID': f"{base_id}-TXT", 'TYPE': 'TEXTBOX', 'DESCRIPTION': 'Remarks'},
                {'TEMPLATEID': f"{base_id}-CHK", 'TYPE': 'CHECKBOX', 'DESCRIPTION': 'Completed'},
                {'TEMPLATEID': f"{base_id}-ACT", 'TYPE': 'LIST', 'DESCRIPTION': 'Corrective Action', 
                 'LOVCODE': config.get('action_lov', '')},
                {'TEMPLATEID': f"{base_id}-DAT", 'TYPE': 'DATE', 'DESCRIPTION': 'Date Completed'},
                {'TEMPLATEID': f"{base_id}-TIM", 'TYPE': 'TIME', 'DESCRIPTION': 'Time Spent'},
                {'TEMPLATEID': f"{base_id}-USR", 'TYPE': 'USER', 'DESCRIPTION': 'Performed By'},
                {'TEMPLATEID': f"{base_id}-SIG", 'TYPE': 'SIGNATURE', 'DESCRIPTION': 'Signature'}
            ]
            
            template_data.extend(entries)
        
        df = pd.DataFrame(template_data)
        df.to_excel(filename, index=False)
    
    def create_formlov(self, filename):
        """Create FORMLOV.xlsx with condition and action values"""
        lov_data = []
        
        for config in self.procedure_lov_config.values():
            condition_values = config['condition_var'].get()
            action_values = config['action_var'].get()
            
            # Add condition LOV entries
            if condition_values and hasattr(config, 'condition_lov'):
                for value in condition_values.split(','):
                    if value.strip():
                        lov_data.append({
                            'LOVCODE': config['condition_lov'],
                            'VALUE': value.strip(),
                            'DESCRIPTION': f"Condition: {value.strip()}",
                            'TYPE': 'CONDITION'
                        })
            
            # Add action LOV entries  
            if action_values and hasattr(config, 'action_lov'):
                for value in action_values.split(','):
                    if value.strip():
                        lov_data.append({
                            'LOVCODE': config['action_lov'],
                            'VALUE': value.strip(),
                            'DESCRIPTION': f"Action: {value.strip()}",
                            'TYPE': 'ACTION'
                        })
        
        df = pd.DataFrame(lov_data)
        df.to_excel(filename, index=False)
    
    def create_formmenu(self, filename):
        """Create FORMMENU.xlsx with simple menu structure"""
        menu_data = [{
            'MENUID': f"{self.sheet_prefix}-MAIN",
            'MENUTEXT': f"Maintenance - {self.selected_sheet}",
            'PARENT': 'ROOT',
            'ORDER': 1,
            'TYPE': 'SECTION',
            'FORM_NAME': f"{self.sheet_prefix}-FORM"
        }]
        
        df = pd.DataFrame(menu_data)
        df.to_excel(filename, index=False)
    
    def save_config(self):
        """Save current configuration"""
        try:
            config_data = {
                'source_file': self.source_file,
                'selected_sheet': self.selected_sheet,
                'user_name': self.get_current_user(),
                'user_mode': self.user_mode.get(),
                'custom_user_entry': self.user_entry.get() if self.user_mode.get() == "custom" else "",
                'file_prefix': getattr(self, 'file_prefix', ''),
                'sheet_prefix': getattr(self, 'sheet_prefix', ''),
                'procedures': [proc['text'] for proc in self.parsed_procedures],
                'lov_configurations': {},
                'timestamp': datetime.now().isoformat()
            }
            
            # Save LOV configurations
            for i, config in self.procedure_lov_config.items():
                config_data['lov_configurations'][i] = {
                    'condition_values': config['condition_var'].get(),
                    'action_values': config['action_var'].get(),
                    'condition_lov': config.get('condition_lov', ''),
                    'action_lov': config.get('action_lov', '')
                }
            
            save_path = filedialog.asksaveasfilename(
                title="Save Configuration",
                defaultextension=".json",
                filetypes=[("JSON files", "*.json")]
            )
            
            if save_path:
                with open(save_path, 'w', encoding='utf-8') as f:
                    json.dump(config_data, f, indent=2, ensure_ascii=False)
                messagebox.showinfo("Configuration Saved", f"Configuration saved to:\n{save_path}")
        
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save configuration: {str(e)}")


def main():
    root = tk.Tk()
    app = ExcelFormatHandler(root)
    
    # Center window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()