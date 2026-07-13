import os
from openpyxl import Workbook
from datetime import datetime

# Create reports directory
os.makedirs("reports", exist_ok=True)

wb = Workbook()

ws = wb.active
ws.title = "Release Summary"

ws["A1"] = "Release Evidence Report"

ws["A3"] = "Repository"
ws["B3"] = os.getenv("GITHUB_REPOSITORY")

ws["A4"] = "Branch"
ws["B4"] = os.getenv("GITHUB_REF_NAME")

ws["A5"] = "Commit SHA"
ws["B5"] = os.getenv("GITHUB_SHA")

ws["A6"] = "Workflow Run ID"
ws["B6"] = os.getenv("GITHUB_RUN_ID")

ws["A7"] = "Triggered By"
ws["B7"] = os.getenv("GITHUB_ACTOR")

ws["A8"] = "Generated On"
ws["B8"] = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")

output = "reports/ReleaseEvidence.xlsx"

wb.save(output)

print(f"Evidence report generated: {output}")
