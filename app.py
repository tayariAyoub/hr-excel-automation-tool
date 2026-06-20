"""Command-line menu for the HR + Excel Automation Tool."""

from __future__ import annotations

from datetime import date

from hr_manager import HRManager


def read_float(prompt: str, minimum: float = 0) -> float:
    """Read a numeric value and keep asking until it is valid."""
    while True:
        try:
            value = float(input(prompt).strip())
            if value < minimum:
                print(f"Please enter a value of at least {minimum}.")
                continue
            return value
        except ValueError:
            print("Please enter a valid number.")


def read_int(prompt: str, minimum: int = 0) -> int:
    """Read an integer value and keep asking until it is valid."""
    while True:
        try:
            value = int(input(prompt).strip())
            if value < minimum:
                print(f"Please enter a value of at least {minimum}.")
                continue
            return value
        except ValueError:
            print("Please enter a whole number.")


def add_employee(manager: HRManager) -> None:
    print("\nAdd employee")
    employee_id = input("Employee ID: ").strip()
    name = input("Name: ").strip()
    department = input("Department: ").strip()
    position = input("Position: ").strip()
    hourly_wage = read_float("Hourly wage: ", minimum=0)
    vacation_days = read_int("Vacation days per year: ", minimum=0)

    manager.add_employee(
        employee_id,
        name,
        department,
        position,
        hourly_wage,
        vacation_days,
    )
    print("Employee added successfully.")


def list_employees(manager: HRManager) -> None:
    employees = manager.list_employees()
    if not employees:
        print("\nNo employees found.")
        return

    print("\nEmployees")
    print("-" * 92)
    print(f"{'ID':<10}{'Name':<24}{'Department':<18}{'Position':<22}{'Wage':>8}")
    print("-" * 92)
    for employee in employees:
        print(
            f"{employee['employee_id']:<10}"
            f"{employee['name']:<24}"
            f"{employee['department']:<18}"
            f"{employee['position']:<22}"
            f"{employee['hourly_wage']:>8.2f}"
        )


def record_hours(manager: HRManager) -> None:
    print("\nRecord working hours")
    work_date = input("Date (YYYY-MM-DD, blank for today): ").strip()
    employee_id = input("Employee ID: ").strip()
    hours = read_float("Hours worked: ", minimum=0)
    notes = input("Notes (optional): ").strip()

    manager.record_working_hours(
        work_date or date.today().isoformat(),
        employee_id,
        hours,
        notes,
    )
    print("Working hours recorded successfully.")


def record_vacation(manager: HRManager) -> None:
    print("\nRecord vacation")
    employee_id = input("Employee ID: ").strip()
    days = read_int("Vacation days used: ", minimum=1)
    manager.record_vacation(employee_id, days)
    print("Vacation balance updated successfully.")


def show_monthly_summary(manager: HRManager) -> None:
    month = input("Month (YYYY-MM): ").strip()
    summary = manager.calculate_monthly_summary(month)

    print(f"\nMonthly summary: {month}")
    print("-" * 80)
    print(f"{'ID':<10}{'Name':<24}{'Hours':>10}{'Wage':>12}{'Salary':>14}")
    print("-" * 80)
    for row in summary:
        print(
            f"{row['employee_id']:<10}"
            f"{row['employee_name']:<24}"
            f"{row['total_hours']:>10.2f}"
            f"{row['hourly_wage']:>12.2f}"
            f"{row['salary']:>14.2f}"
        )


def export_report(manager: HRManager) -> None:
    month = input("Month to export (YYYY-MM): ").strip()
    report_path = manager.export_monthly_report(month)
    print(f"Report exported to: {report_path}")


def main() -> None:
    manager = HRManager()
    actions = {
        "1": add_employee,
        "2": list_employees,
        "3": record_hours,
        "4": record_vacation,
        "5": show_monthly_summary,
        "6": export_report,
    }

    while True:
        print(
            "\nHR + Excel Automation Tool\n"
            "1. Add employee\n"
            "2. List employees\n"
            "3. Record daily working hours\n"
            "4. Record vacation days\n"
            "5. Calculate monthly salary summary\n"
            "6. Export monthly HR report to Excel\n"
            "0. Exit"
        )
        choice = input("Choose an option: ").strip()

        if choice == "0":
            print("Goodbye!")
            break

        action = actions.get(choice)
        if action is None:
            print("Invalid option. Please choose a number from the menu.")
            continue

        try:
            action(manager)
        except ValueError as error:
            print(f"Error: {error}")
        except OSError as error:
            print(f"Could not access an Excel file: {error}")


if __name__ == "__main__":
    main()
