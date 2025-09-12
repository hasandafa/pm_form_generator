# Excel Format Handler v3.0 - Maintenance Form Generator

> Advanced Excel maintenance tasklist processor with intelligent format detection and automated form generation
> 
> **Author:** MK.ABDULLAH.DAFA (hasandafa)  
> **Version:** 3.0 - Focused Solution  
> **Last Updated:** September 2025

## 🚀 Overview

Excel Format Handler v3.0 is a comprehensive Python application designed to process Excel-based maintenance tasklists and convert them into structured database forms. The application features intelligent format detection, automated procedure extraction, and streamlined LOV (List of Values) configuration for efficient maintenance form generation.

### Key Improvements in v3.0
- **Simplified Interface**: Clean 4-step workflow instead of complex multiple tabs
- **Smart Detection**: Advanced pattern recognition for various Excel formats
- **Real-time LOV Generation**: Automatic unique code generation as you type
- **Conflict Resolution**: Built-in unique identifier system to prevent duplicates
- **Enhanced User Experience**: Default user settings with custom override options

---

## 📋 Requirements

### System Requirements
- **Operating System**: Windows 7/8/10/11, macOS 10.14+, or Linux Ubuntu 18.04+
- **Python**: Version 3.8 or higher
- **Memory**: Minimum 4GB RAM (8GB recommended for large Excel files)
- **Storage**: 500MB free space for installation and temporary files

### Python Dependencies
```
pandas >= 1.5.0          # Excel file processing and data manipulation
openpyxl >= 3.1.0        # Excel file reading/writing support
tkinter                  # GUI framework (included with Python)
pathlib                 # File path operations (included with Python 3.4+)
hashlib                 # Unique identifier generation (standard library)
```

---

## 🔧 Installation Guide

### Method 1: Run from Source (Recommended for Developers)

1. **Install Python Dependencies**
   ```bash
   # Create virtual environment (recommended)
   python -m venv excel_handler_env
   
   # Activate virtual environment
   # Windows:
   excel_handler_env\Scripts\activate
   # macOS/Linux:
   source excel_handler_env/bin/activate
   
   # Install required packages
   pip install pandas>=1.5.0 openpyxl>=3.1.0
   ```

2. **Download and Run Application**
   ```bash
   # Clone or download the project
   git clone https://github.com/hasandafa/pm_form_generator.git
   cd pm_form_generator
   
   # Run the application
   python formgenerator.py
   ```

### Method 2: Build Standalone Executable

1. **Install PyInstaller**
   ```bash
   pip install pyinstaller
   ```

2. **Build Executable**
   ```bash
   # Create standalone executable
   pyinstaller --onefile --windowed formgenerator.py
   
   # Run the executable
   dist/formgenerator.exe
   ```

### Method 3: Quick Setup Script

Create a `setup.bat` file:
```batch
@echo off
echo Installing Excel Format Handler v3.0...
python -m pip install --upgrade pip
pip install pandas>=1.5.0 openpyxl>=3.1.0
echo Installation complete!
echo Run: python formgenerator.py
pause
```

---

## 📖 How to Use - Step by Step Guide

### Step 1: File Selection

1. **Launch Application**
   ```bash
   python formgenerator.py
   ```

2. **Select Excel File**
   - Click **"Select Excel File"** button
   - Browse to your maintenance tasklist Excel file
   - Supported formats: `.xlsx`, `.xls`
   - File automatically loads and displays available sheets

3. **Configure User Settings**
   - **Default User**: "MK.ABDULLAH.DAFA" (recommended)
   - **Custom User**: Select "Custom" radio button and enter your name
   - This will appear in all generated form metadata

4. **Choose Sheet**
   - Select the specific worksheet containing your maintenance procedures
   - Sheet names appear in the dropdown menu
   - Click **"Analyze Format"** to proceed

**Expected Result**: File loads successfully, unique prefixes generated automatically
```
Generated Prefixes: File: MAI-A3F | Sheet: ENG-B2C
```

### Step 2: Format Analysis

1. **Automatic Format Detection**
   - Application analyzes the selected sheet content
   - Detects format type (Maintenance, Checklist, Calibration, etc.)
   - Calculates confidence score for detection accuracy
   - Displays detected procedures with row/column information

2. **Review Analysis Results**
   ```
   📋 FORMAT ANALYSIS RESULTS
   ==================================================
   📁 File: YKN-CPP2-G-603_PM1.xlsx
   📄 Sheet: ENGINE MAINTENANCE
   🔍 Detected Format: MAINTENANCE
   🎯 Confidence: 75.0%
   
   ✅ Found 15 procedures:
   ----------------------------------------
    1. 1. Inspect engine oil level and condition...
    2. 2. Check cooling system for leaks...
    3. 3. Test engine temperature sensors...
   ```

3. **Action Options**
   - **Accept & Configure LOVs**: Proceed with detected procedures (recommended)
   - **Re-analyze**: Run detection again if results seem incorrect
   - **View Raw Data**: Examine raw Excel structure for manual verification

**Expected Result**: Clear analysis showing detected procedures and confidence level

### Step 3: LOV Configuration

1. **Automatic LOV Setup**
   - Click **"Accept & Configure LOVs"**
   - Application creates input fields for each detected procedure
   - Scrollable interface accommodates large numbers of procedures

2. **Configure Individual Procedures**
   
   For each procedure row:
   - **Procedure**: Displays the detected maintenance procedure text
   - **Condition Values**: Enter possible condition states (comma-separated)
   - **Action Values**: Enter possible corrective actions (comma-separated)
   - **Generated LOV Code**: Automatically updates as you type values

   **Example Configuration:**
   ```
   Procedure: Inspect engine oil level and condition
   Condition Values: Good,Dirty,Low,Contaminated
   Action Values: No Action,Top Up,Change Oil,Clean System
   Generated LOV Code: C:ENG-B2C-GDLC | A:ENG-B2C-NTCC
   ```

3. **Bulk Configuration Options**
   - **Auto-Configure Common LOVs**: Automatically assigns standard values based on procedure keywords
     - "inspect" → Good,Dirty,Damaged,Missing / No Action,Clean,Repair,Replace
     - "check" → OK,Not OK,Needs Attention / No Action,Adjust,Repair
     - "clean" → Clean,Dirty,Blocked / Cleaned,Replaced
     - "calibrate" → In Tolerance,Out of Tolerance / Calibrated,Adjusted

   - **Clear All LOVs**: Removes all configured values to start fresh

**Expected Result**: Each procedure has configured condition/action values with unique LOV codes

### Step 4: Generate Forms

1. **Set Output Directory**
   - Specify where generated files will be saved
   - Default: Current working directory
   - Click **"Browse"** to select different location

2. **Preview Generation** (Optional but Recommended)
   - Click **"Preview Generation"** to review what will be created
   - Shows file information, user settings, and configuration status
   ```
   FORM GENERATION PREVIEW
   ==================================================
   Source File: YKN-CPP2-G-603_PM1.xlsx
   Sheet: ENGINE MAINTENANCE
   User: MK.ABDULLAH.DAFA
   File Prefix: YKN-A3F
   Sheet Prefix: ENG-B2C
   
   FILES TO BE GENERATED:
   ✓ FORMHEAD.xlsx - Form metadata and configuration
   ✓ FORMTEMPLATE.xlsx - 135 template entries
   ✓ FORMLOV.xlsx - LOV definitions
   ✓ FORMMENU.xlsx - Menu structure
   
   CONFIGURATION STATUS:
   ✓ Configured: 15/15
   ```

3. **Generate Excel Forms**
   - Click **"Generate Excel Forms"** to create output files
   - Progress indication during generation process
   - Success dialog shows created files and output directory

4. **Generated Files Structure**
   ```
   Output Directory/
   ├── FORMHEAD_ENG-B2C_20250913_143022.xlsx     # Form metadata
   ├── FORMTEMPLATE_ENG-B2C_20250913_143022.xlsx # 9 entries per procedure
   ├── FORMLOV_ENG-B2C_20250913_143022.xlsx      # Condition/Action values
   └── FORMMENU_ENG-B2C_20250913_143022.xlsx     # Menu structure
   ```

**Expected Result**: 4 Excel files generated with unique timestamps and prefixes

---

## 📁 Input File Format Requirements

### Excel File Structure
```
Required Elements:
✓ Excel file (.xlsx or .xls format)
✓ At least one worksheet with maintenance procedures
✓ Procedures should be numbered (1., 2., 3., etc.) or contain action keywords
✓ Text length minimum 10 characters per procedure

Optional Elements:
• Section headers (ENGINE, GENERATOR, ELECTRICAL, etc.)
• Additional columns with condition/action data
• Comments or remarks columns
```

### Supported Format Patterns

1. **Standard Maintenance Format**
   ```
   1. Inspect engine oil level and condition
   2. Check cooling system for leaks and damage
   3. Test all engine sensors and connections
   ```

2. **Checklist Format**
   ```
   Check startup parameters    OK / Not OK
   Verify system pressure     Normal / High / Low
   Monitor temperature        Within Range / Alert
   ```

3. **Calibration Format**
   ```
   Transmitter TT-001    As Found: 95.2°C    As Left: 95.5°C
   Pressure PT-002       As Found: 1.85 bar  As Left: 1.87 bar
   ```

### Example Compatible Files
- YKN-CPP2-G-603_PM1.xlsx (Gas Turbine Maintenance)
- ELECTRICAL-INSP-2024.xlsx (Electrical Inspection)
- CALIBRATION-INST-Q3.xlsx (Instrument Calibration)

---

## 📊 Output Files Specification

### 1. FORMHEAD.xlsx - Form Metadata
```
Columns: FORMNAME, TITLE, DESCRIPTION, CREATED_BY, CREATED_DATE, 
         SOURCE_FILE, SOURCE_SHEET, PROCEDURES_COUNT, FILE_PREFIX, SHEET_PREFIX

Example Row:
ENG-B2C-FORM | Maintenance Form - ENGINE MAINTENANCE | Generated from YKN-CPP2-G-603_PM1.xlsx | 
MK.ABDULLAH.DAFA | 2025-09-13 14:30:22 | YKN-CPP2-G-603_PM1.xlsx | ENGINE MAINTENANCE | 
15 | YKN-A3F | ENG-B2C
```

### 2. FORMTEMPLATE.xlsx - Field Definitions (9 entries per procedure)
```
Columns: TEMPLATEID, TYPE, DESCRIPTION, LOVCODE

Example Entries for Procedure 1:
ENG-B2C-P001-LBL | LABEL     | Inspect engine oil level and condition | 
ENG-B2C-P001-LST | LIST      | Condition Found                        | ENG-B2C-GDLC
ENG-B2C-P001-TXT | TEXTBOX   | Remarks                               |
ENG-B2C-P001-CHK | CHECKBOX  | Completed                             |
ENG-B2C-P001-ACT | LIST      | Corrective Action                     | ENG-B2C-NTCC
ENG-B2C-P001-DAT | DATE      | Date Completed                        |
ENG-B2C-P001-TIM | TIME      | Time Spent                            |
ENG-B2C-P001-USR | USER      | Performed By                          |
ENG-B2C-P001-SIG | SIGNATURE | Signature                             |
```

### 3. FORMLOV.xlsx - List of Values
```
Columns: LOVCODE, VALUE, DESCRIPTION, TYPE

Example Entries:
ENG-B2C-GDLC | Good         | Condition: Good         | CONDITION
ENG-B2C-GDLC | Dirty        | Condition: Dirty        | CONDITION
ENG-B2C-GDLC | Low          | Condition: Low          | CONDITION
ENG-B2C-GDLC | Contaminated | Condition: Contaminated | CONDITION
ENG-B2C-NTCC | No Action    | Action: No Action       | ACTION
ENG-B2C-NTCC | Top Up       | Action: Top Up          | ACTION
ENG-B2C-NTCC | Change Oil   | Action: Change Oil      | ACTION
ENG-B2C-NTCC | Clean System | Action: Clean System    | ACTION
```

### 4. FORMMENU.xlsx - Menu Structure
```
Columns: MENUID, MENUTEXT, PARENT, ORDER, TYPE, FORM_NAME

Example Entry:
ENG-B2C-MAIN | Maintenance - ENGINE MAINTENANCE | ROOT | 1 | SECTION | ENG-B2C-FORM
```

---

## 🛠 Troubleshooting Guide

### Common Issues and Solutions

#### Issue: "No procedures detected"
**Cause**: Excel format not recognized or procedures too short
**Solutions**:
1. Check that procedures are numbered (1., 2., etc.)
2. Ensure procedure text is at least 10 characters long
3. Use "View Raw Data" to examine structure
4. Try manual format override in Manual Override tab

#### Issue: "AttributeError: 'ExcelFormatHandler' object has no attribute..."
**Cause**: Code version mismatch or missing methods
**Solution**: Ensure you're using the complete v3.0 code with all methods implemented

#### Issue: LOV codes not generating
**Cause**: Empty values or special characters in input
**Solutions**:
1. Ensure condition/action values are comma-separated
2. Avoid special characters in values
3. Check that values are not empty

#### Issue: Generated files have duplicate identifiers
**Cause**: Unique identifier system not working
**Solution**: Delete `unique_identifiers.json` file and restart application

#### Issue: Excel file won't load
**Cause**: Corrupted file or unsupported format
**Solutions**:
1. Verify file is .xlsx or .xls format
2. Try opening file in Excel first to check for corruption
3. Ensure file is not password protected
4. Check file permissions

### Performance Optimization

**For Large Files (>1000 procedures):**
1. Close other applications to free memory
2. Use "Auto-Configure Common LOVs" instead of manual configuration
3. Consider splitting large files into smaller sheets
4. Generate files one at a time if memory issues occur

**Network/Enterprise Environments:**
1. Ensure write permissions to output directory
2. Check antivirus software isn't blocking file operations
3. Consider running as administrator if file access issues persist

---

## 🔍 Advanced Features

### Unique Identifier System
- Automatic generation of file and sheet prefixes
- Collision detection and resolution
- Persistent tracking in `unique_identifiers.json`
- Guaranteed uniqueness across all generated forms

### Smart LOV Code Generation
```
Input: "Good,Damaged,Missing" 
Output: "SHEET_PREFIX-GDM"

Input: "Good,Damaged,Missing" (duplicate)
Output: "SHEET_PREFIX-GDM1" (auto-incremented)
```

### Configuration Persistence
- Save current configuration to JSON file
- Load previous configurations
- Session state preservation
- User preference storage

### Batch Processing Capabilities
- Process multiple Excel files sequentially
- Consistent LOV assignment across files
- Bulk configuration options
- Progress tracking and error handling

---

## 📝 Best Practices

### File Naming Conventions
```
Input Files:  FACILITY-EQUIPMENT-TYPE-SEQUENCE.xlsx
Example:      YKN-CPP2-G-603_PM1.xlsx

Output Files: FORMTYPE_PREFIX_YYYYMMDD_HHMMSS.xlsx
Example:      FORMHEAD_ENG-B2C_20250913_143022.xlsx
```

### LOV Configuration Guidelines
1. **Keep values concise**: Use short, clear descriptive terms
2. **Be consistent**: Use similar value sets for similar procedures
3. **Cover all scenarios**: Include "No Action" for actions, "Good" for conditions
4. **Avoid duplicates**: Each value should represent a distinct state or action

### Quality Assurance Checklist
- [ ] All procedures have configured LOVs
- [ ] Generated LOV codes are unique
- [ ] Preview shows correct procedure count
- [ ] Output directory has write permissions
- [ ] User name is correctly configured

---

## 🤝 Git Workflow Integration

This project follows professional Git workflow practices for multi-device development:

### Repository Management
```bash
# Daily workflow
git pull origin main           # Get latest changes
git add .                      # Stage changes
git commit -m "descriptive message"  # Commit changes
git push origin main           # Push to GitHub

# Branch management for features
git checkout -b feature-new-parser
git commit -m "Add support for new format"
git checkout main
git merge feature-new-parser
```

### Project Structure
```
pm_form_generator/
├── formgenerator.py           # Main application
├── requirements.txt           # Python dependencies  
├── README.md                  # This documentation
├── ui.html                    # Visual workflow guide
├── build.bat                  # Build script for executable
├── unique_identifiers.json   # Unique ID tracking (auto-generated)
└── format_learning_db.json   # Learning system data (auto-generated)
```

---

## 📞 Support and Contributing

### Getting Help
1. **Check this README** for common issues and solutions
2. **Review the troubleshooting section** for specific error messages
3. **Use "View Raw Data"** feature to diagnose format detection issues
4. **Check GitHub Issues** for known problems and solutions

### Contributing
1. Fork the repository
2. Create a feature branch (`git checkout -b feature-amazing-feature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature-amazing-feature`)
5. Open a Pull Request

### Version History
- **v3.0**: Complete redesign with simplified workflow and enhanced reliability
- **v2.4**: Multi-tab interface with advanced detection
- **v2.0**: Basic Excel processing with manual LOV configuration
- **v1.0**: Initial release with simple form generation

---

## 📄 License

This project is provided as-is for educational and practical purposes. Feel free to modify and distribute according to your needs.

**Developed with ❤️ by MK.ABDULLAH.DAFA**

---

**Happy Coding! 🚀**