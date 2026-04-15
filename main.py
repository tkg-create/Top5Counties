import pandas as pd

file_path = "qcew_full_report.xlsx"

xls = pd.ExcelFile(file_path)

all_data = []

# Load each county sheet
for sheet in xls.sheet_names:

    df = pd.read_excel(
        file_path,
        sheet_name=sheet,
        skiprows=7,
        header=None
    )

    df["county"] = sheet

    all_data.append(df)

data = pd.concat(all_data, ignore_index=True)
data = data[data["county"] != "Pennsylvania"]  # Prevent contamination from state-level statistics


# Rename columns
data = data.rename(columns={
    0: "naics",
    1: "industry",
    2: "establishments",
    3: "employment",
    4: "wage",
    5: "employment_change",
    6: "wage_change",
    7: "employment_growth",
    8: "wage_growth"
})

# Drop junk column
if 9 in data.columns:
    data = data.drop(columns=[9])

print(data.columns)