"""Builds a cost-summary CSV for every model in matches.csv.

For each matched BOM file, reads the summary block at the bottom of its
'Costed Material BOM' sheet (labels in column A, values in column C) and
writes one row per model+section in long format:
Job Number, Model, Filename, Section, Value.
"""
import csv
import openpyxl

LABELS = [
    "Material",
    "Material Burden",
    "Labor & OH",
    "Assembly",
    "Total Cost",
    "Total Selling Price (less freight, tax, etc)",
    "Gross Margin %",
]

OUT_FILE = "cost_summary.csv"


def read_summary(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb["Costed Material BOM"]
        values = {}
        for row in ws.iter_rows(values_only=True):
            label = str(row[0]).strip() if row[0] is not None else ""
            if label in LABELS:
                values[label] = row[2]
        return values
    finally:
        wb.close()


def main():
    with open("matches.csv", newline="") as f:
        matches = list(csv.DictReader(f))

    out_rows = []
    problems = []
    for match in matches:
        file_name = match["file"]
        job_number = file_name.split("_")[0]
        try:
            summary = read_summary(rf"..\BOMs\{file_name}")
        except Exception as e:
            problems.append(f"{file_name}: {e}")
            continue
        missing = [l for l in LABELS if l not in summary]
        if missing:
            problems.append(f"{file_name}: missing labels {missing}")
        for label in LABELS:
            out_rows.append(
                [job_number, match["part_number"], file_name, label, summary.get(label, "")]
            )

    with open(OUT_FILE, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Job Number", "Model", "Filename", "Section", "Value"])
        writer.writerows(out_rows)

    print(f"Wrote {len(out_rows)} rows to {OUT_FILE}")
    if problems:
        print(f"{len(problems)} problems:")
        for p in problems:
            print("  ", p)


if __name__ == "__main__":
    main()
