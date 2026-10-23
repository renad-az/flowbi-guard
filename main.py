from pathlib import Path
from openpyxl import load_workbook

# تحديد مكان المشروع وملف المصدر
project_folder = Path(__file__).resolve().parent

source_file = (
    project_folder
    / "data"
    / "incoming"
    / "FlowBI_Project_Management.xlsx"
)

# فتح ملف Excel للقراءة فقط
workbook = load_workbook(source_file, read_only=True)

try:
    print("Source:", source_file)
    print("Sheets:", workbook.sheetnames)

    # الشيتات التي يعتمد عليها الداشبورد
    required_sheets = [
        "Projects",
        "WeeklySnapshots",
        "Risks",
        "Calendar",
    ]

    # البحث عن الشيتات الناقصة
    missing_sheets = [
        name
        for name in required_sheets
        if name not in workbook.sheetnames
    ]

    if missing_sheets:
        print("BLOCKED: Missing required sheets:", missing_sheets)
    else:
        print("PASS: All required sheets are present.")

finally:
    # إغلاق الملف حتى لو حدث خطأ أثناء الفحص
    workbook.close()