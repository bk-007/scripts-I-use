# Scripts I Use 

This repository contains scripts I use for work and home.

## Scripts

### 1. ems-gen.py

**Purpose**: Generates a sample Employee Management System Excel workbook with realistic data.

**Features**:
- Creates a multi-sheet Excel workbook with sample data
- Includes tables for employees, training processes, training status, one-on-ones, projects, and onboarding tasks
- Generates realistic relationships between data across sheets
- Uses Faker library to create realistic fake data

**Usage**:
```bash
python ems-gen.py [output_file.xlsx]
```

### 1. ems-gen-up.py

**Purpose**: Provides a GUI interface to update and manage Employee Management System Excel workbooks.

**Features**:
- Tkinter-based GUI for easy interaction
- View and edit all sheets in the workbook
- Add, edit, and delete records
- Save changes back to the Excel file
- Preserves Excel formatting and formulas

**Usage**:
```bash
python ems-gen-up.py [output_file.xlsx]
