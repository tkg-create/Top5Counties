import pandas as pd
from sklearn.preprocessing import MinMaxScaler
from sklearn.ensemble import RandomForestRegressor

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

# print(data.columns)


# Clean numeric values
def clean_number(x):
    if isinstance(x, str):
        x = x.replace(",", "")
    return x


def clean_percent(x):
    if isinstance(x, str):
        x = x.replace("%", "")
    return x


data["employment"] = data["employment"].apply(clean_number)
data["wage"] = data["wage"].apply(clean_number)

data["employment_growth"] = data["employment_growth"].apply(clean_percent)

# Convert to numeric
data["employment"] = pd.to_numeric(data["employment"], errors="coerce")
data["wage"] = pd.to_numeric(data["wage"], errors="coerce")
data["employment_growth"] = pd.to_numeric(data["employment_growth"], errors="coerce")

# Remove missing rows
data = data.dropna(subset=["industry", "employment", "wage", "employment_growth"])

# Remove total rows
data = data[data["industry"] != "Total"]

# Generate industry list
industries = sorted(data["industry"].unique())

print("\nAvailable industries:\n")

for i, ind in enumerate(industries):
    print(f"{i+1}. {ind}")

# Get user selected industry
choice = int(input("\nSelect industry number: "))
selected_industry = industries[choice-1]

print("\nSelected:", selected_industry)

# Filter dataset
industry_data = data[data["industry"] == selected_industry].copy()

# Normalize features
features = industry_data[["employment", "wage", "employment_growth"]]

scaler = MinMaxScaler()
scaled = scaler.fit_transform(features)

industry_data["emp_score"] = scaled[:, 0]
industry_data["wage_score"] = scaled[:, 1]
industry_data["growth_score"] = scaled[:, 2]

# Create Opportunity Index
industry_data["opportunity_index"] = (
    0.5 * industry_data["emp_score"] +
    0.3 * industry_data["growth_score"] +
    0.2 * industry_data["wage_score"]
)

# Train Random Forest
X = industry_data[["emp_score", "growth_score", "wage_score"]]
y = industry_data["opportunity_index"]

model = RandomForestRegressor(
    n_estimators=200,
    random_state=42
)

model.fit(X, y)

# Predict county scores
industry_data["predicted_score"] = model.predict(X)

# Rank counties
ranked = industry_data.sort_values(
    by="predicted_score",
    ascending=False
)

top5 = ranked[["county", "predicted_score"]].head(5)

# Output
print("\nTop Counties for", selected_industry)

for i, row in enumerate(top5.itertuples(), start=1):
    print(f"{i}. {row.county} (Score: {row.predicted_score:.3f})")
