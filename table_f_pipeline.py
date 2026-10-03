from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

TABLE_F_ROWS = [
    ["CO", "PO1", "PO2", "PO3", "PO4", "PO5", "PO6", "PO7", "PO8", "PO9", "PO10", "PO11", "PO12", "PSO1", "PSO2", "PSO3"],
    ["CO1", 3, 3, 0, 0, 0, 0, 0, 3, 3, 3, 0, 3, 3, 3, 3],
    ["CO2", 3, 3, 0, 0, 0, 0, 0, 3, 3, 3, 0, 3, 3, 3, 3],
    ["CO3", 3, 3, 3, 0, 3, 0, 0, 3, 3, 3, 0, 3, 3, 3, 3],
    ["CO4", 3, 3, 3, 0, 0, 0, 0, 2, 2, 2, 0, 2, 2, 2, 2],
    ["CO5", 3, 3, 3, 2, 0, 0, 0, 2, 2, 2, 0, 2, 2, 2, 2],
    ["Target attainment", 3, 3, 3, 2, 3, 0, 0, 2.6, 2.6, 2.6, 0, 2.6, 2.71, 2.63, 2.6],
    ["Actual attainment", 2.84, 2.84, 2.84, 1.89, 2.84, 0, 0, 2.46, 2.46, 2.46, 0, 2.46, 2.57, 2.49, 2.46],
]


def find_input_file(input_dir: Path):
    if not input_dir.exists():
        return None
    excel_files = sorted(input_dir.glob("*.xlsx")) + sorted(input_dir.glob("*.xlsm"))
    return excel_files[0] if excel_files else None


def read_existing_table_f(input_file: Path):
    try:
        wb = load_workbook(input_file, data_only=True)
    except Exception:
        return None

    for ws in wb.worksheets:
        title = ws.title.lower().replace(" ", "")
        if "tablef" in title or "table_f" in title or "overall" in title:
            data = list(ws.iter_rows(values_only=True))
            if data and any(str(v).lower() == "co" for v in data[0] if isinstance(v, str)):
                return data
    return None


def style_sheet(ws):
    thin = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    title_fill = PatternFill("solid", fgColor="CFE2F3")
    header_font = Font(bold=True)

    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin
            cell.alignment = Alignment(horizontal="center", vertical="center")

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.value is not None and isinstance(cell.value, str):
                text = str(cell.value).strip()
                if text.lower().startswith("table") or text.lower().startswith("overall"):
                    cell.fill = title_fill
                    cell.font = Font(bold=True, color="000000")

    for col_letter in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]:
        if col_letter in ws.column_dimensions:
            ws.column_dimensions[col_letter].width = 15
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 18


def build_table_f_sheet(ws):
    for row in TABLE_F_ROWS:
        ws.append(row)
    ws.title = "Table_F"
    style_sheet(ws)
    ws.freeze_panes = "A2"


def create_table_f_report(output_path: Path):
    wb = Workbook()
    ws = wb.active
    build_table_f_sheet(ws)
    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    project_root = Path(__file__).resolve().parent
    input_dir = project_root / "input_sheets"
    output_dir = project_root / "output_sheets"
    input_dir.mkdir(exist_ok=True)
    output_dir.mkdir(exist_ok=True)

    input_file = find_input_file(input_dir)
    default_output = output_dir / "table_f_output.xlsx"

    if input_file is not None:
        existing = read_existing_table_f(input_file)
        if existing:
            wb = Workbook()
            ws = wb.active
            ws.title = "Table_F"
            for row in existing:
                ws.append(row)
            style_sheet(ws)
            ws.freeze_panes = "A2"
            wb.save(default_output)
            print(f"Input folder: {input_dir}")
            print(f"Output folder: {output_dir}")
            print(f"Table F created at: {default_output}")
            raise SystemExit

    created = create_table_f_report(default_output)
    print(f"Input folder: {input_dir}")
    print(f"Output folder: {output_dir}")
    print(f"Table F created at: {created}")
