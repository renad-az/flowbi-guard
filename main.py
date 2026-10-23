from pathlib import Path
from openpyxl import load_workbook

project_folder = Path(__file__).resolve().parent

source_file = (
    project_folder
    / "data"
    / "incoming"
    / "FlowBI_Project_Management.xlsx"
)

required_columns = {
    "Projects": [
        "ProjectID", "ProjectName", "Department",
        "ProjectManager", "City", "Program", "Priority",
        "StartDate", "PlannedEndDate", "BudgetSAR",
        "AsOfDate", "ProjectStatus",
    ],
    "WeeklySnapshots": [
        "ProjectID", "SnapshotDate", "PlannedProgress",
        "ActualProgress", "PlannedValueSAR",
        "EarnedValueSAR", "ActualCostSAR", "PeriodCostSAR",
        "EstimateToCompleteSAR", "ScheduleDelayDays",
        "ScenarioNote",
    ],
    "Risks": [
        "RiskID", "ProjectID", "RiskDescription", "Category",
        "ProbabilityScore", "ImpactScore", "Owner",
        "OpenedDate", "ClosedDate", "RiskStatus",
        "Mitigation", "TargetActionDate",
    ],
    "Calendar": [
        "Date", "Year", "MonthNumber", "YearMonth", "Quarter",
    ],
}

workbook = load_workbook(source_file, read_only=True)

try:
    errors = []

    # فحص الشيتات وأسماء الأعمدة
    for sheet_name, columns in required_columns.items():
        if sheet_name not in workbook.sheetnames:
            errors.append(f"Missing sheet: {sheet_name}")
            continue

        sheet = workbook[sheet_name]

        headers = next(
            sheet.iter_rows(
                min_row=1,
                max_row=1,
                values_only=True,
            ),
            (),
        )

        for column_name in columns:
            if column_name not in headers:
                errors.append(
                    f"{sheet_name}: Missing column '{column_name}' "
                    "(header row 1)"
                )

    # نفحص القيم فقط إذا كانت بنية الملف سليمة
    if not errors:
        sheet = workbook["Projects"]
        rows = sheet.iter_rows(values_only=True)
        headers = next(rows)

        project_id_index = headers.index("ProjectID")
        seen_ids = {}

        for row_number, row in enumerate(rows, start=2):

            # تجاهل الصفوف الفارغة بالكامل
            if all(value is None for value in row):
                continue

            project_id = row[project_id_index]

            if (
                project_id is None
                or str(project_id).strip() == ""
            ):
                errors.append(
                    f"Projects: Missing ProjectID "
                    f"(row {row_number})"
                )
                continue

            if not isinstance(project_id, str):
                errors.append(
                    f"Projects: ProjectID must be text "
                    f"(row {row_number})"
                )
                continue

            # نستخدم مفتاح مقارنة بدون مسافات وبأحرف موحدة
            comparison_id = project_id.strip().upper()

            if comparison_id in seen_ids:
                first_row = seen_ids[comparison_id]

                errors.append(
                    f"Projects: Duplicate ProjectID "
                    f"'{project_id}' (row {row_number}; "
                    f"first found at row {first_row})"
                )
            else:
                seen_ids[comparison_id] = row_number

    if errors:
        print("BLOCKED: Validation failed.")

        for error in errors:
            print("-", error)

    else:
        print(
            "PASS: Required sheets, columns, "
            "and project IDs are valid."
        )

finally:
    workbook.close()