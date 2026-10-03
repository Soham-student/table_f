# Table F Pipeline

This folder is created to generate only the final Table F output.

## Folders

- input_sheets: place your Excel file here
- output_sheets: generated output is saved here

## Run

From this folder, run:

```powershell
python .\table_f_pipeline.py
```

## Output

The script creates an Excel file named:

```text
output_sheets\table_f_output.xlsx
```

This workbook contains the Table F sheet with the CO, PO, PSO, target attainment, and actual attainment data.
