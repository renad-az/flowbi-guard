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

    if errors:
        print("BLOCKED: Workbook structure validation failed.")

        for error in errors:
            print("-", error)
    else:
        print("PASS: All required sheets and columns are present.")

finally:
    workbook.close()