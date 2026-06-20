"""Create realistic demonstration data for the portfolio project."""

from hr_manager import HRManager


def main() -> None:
    manager = HRManager()

    if not manager.list_employees():
        manager.add_employee("E001", "Anna Keller", "Engineering", "Software Developer", 28.50, 30)
        manager.add_employee("E002", "David Weber", "Sales", "Sales Assistant", 22.00, 28)
        manager.add_employee("E003", "Mira Hassan", "Operations", "Office Coordinator", 24.50, 30)

        sample_hours = [
            ("2026-06-01", "E001", 8, "Backend development"),
            ("2026-06-02", "E001", 7.5, "Testing and documentation"),
            ("2026-06-01", "E002", 8, "Customer calls"),
            ("2026-06-02", "E002", 8, "Sales administration"),
            ("2026-06-01", "E003", 6.5, "Office coordination"),
            ("2026-06-02", "E003", 7, "Supplier management"),
        ]
        for work_date, employee_id, hours, notes in sample_hours:
            manager.record_working_hours(work_date, employee_id, hours, notes)

        manager.record_vacation("E002", 2)

    report = manager.export_monthly_report("2026-06")
    print(f"Sample files created in: {manager.data_dir.resolve()}")
    print(f"Sample monthly report: {report}")


if __name__ == "__main__":
    main()
