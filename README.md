# HR + Excel Automation Tool

A beginner-friendly Python application that helps small companies manage
employee information, working hours, vacation balances, salaries, and monthly
HR reports using familiar Excel files.

The project is designed as a practical portfolio example for Python automation,
Fiverr, Upwork, internships, and working-student applications.

## Features

- Add and list employees
- Record daily working hours
- Calculate monthly hours and salary
- Track used and remaining vacation days
- Export a formatted monthly HR report to Excel
- Create all required Excel files automatically
- Validate dates, employee IDs, hours, wages, and vacation balances
- Include realistic sample data and automated tests

## Project structure

```text
hr-excel-automation/
|-- app.py                 # Interactive command-line menu
|-- hr_manager.py          # HR logic and Excel operations
|-- create_sample_data.py  # Creates demonstration employees and reports
|-- requirements.txt
|-- tests/
|   `-- test_hr_manager.py
`-- data/                  # Excel files are generated here
```

## Excel files

The application automatically creates:

- `employees.xlsx`
- `working_hours.xlsx`
- `monthly_report.xlsx`

The workbooks include formatted headers, filters, frozen header rows, readable
column widths, and appropriate date and EUR currency formats.

## Installation

1. Install Python 3.10 or newer.
2. Download or clone this repository.
3. Open a terminal in the project folder.
4. Install the dependency:

```bash
pip install -r requirements.txt
```

## How to run

Start the interactive menu:

```bash
python app.py
```

Create realistic sample data and a June 2026 report:

```bash
python create_sample_data.py
```

Run the automated tests:

```bash
python -m unittest discover -s tests -v
```

## Example company use case

A small consulting company employs several hourly workers. The office manager
can add each employee once, record daily hours, update vacation use, and export
a monthly Excel report showing:

- Total hours per employee
- Hourly wage
- Calculated monthly salary
- Vacation days used
- Vacation days remaining

The exported report can be reviewed by management or passed to payroll.

## Important note

This is an educational portfolio project, not a replacement for certified
payroll, tax, or HR software. Real companies should also consider access
controls, backups, privacy requirements, taxes, overtime rules, and labor law.

## Possible future improvements

- Graphical desktop interface
- Employee search and editing
- Overtime and tax calculations
- Password-protected workbooks
- Database storage for larger companies

## Author

Ayoub Tayari  
Computer Engineering student at RWTH Aachen University
