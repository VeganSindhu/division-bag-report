import pandas as pd
from pathlib import Path

print("🔄 Starting Division-wise Report Generator...")

# Reference file (fixed)
reference_filename = "division wis.xlsx"
reference_path = Path(reference_filename)

if not reference_path.exists():
    print(f"❌ Reference file '{reference_filename}' not found.")
    exit()

# Load reference data
reference_df = pd.read_excel(reference_path)
reference_df.columns = reference_df.columns.str.strip().str.lower()
print(f"✅ Loaded reference file: {reference_filename}")

# Collect all input files (.csv, .xls, .xlsx)
input_files = [
    f for f in Path(".").glob("*.*")
    if f.suffix.lower() in [".csv", ".xls", ".xlsx"] and f.name != reference_filename
]

if not input_files:
    print("⚠️ No input Excel files found in this folder.")
    exit()

print(f"📂 Found {len(input_files)} input files:")
for f in input_files:
    print(f"   • {f.name}")

# Combine all input data
all_data = []

for file in input_files:
    try:
        # Read file
        if file.suffix.lower() == ".csv":
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)

        df.columns = df.columns.str.strip().str.lower()

        # Determine bag type based on filename
        fname = file.name.lower()
        if "set1" in fname:
            df["bag type"] = "PL"
        elif "set2" in fname:
            df["bag type"] = "SP"

        all_data.append(df)
    except Exception as e:
        print(f"❌ Error reading {file.name}: {e}")

if not all_data:
    print("⚠️ No valid data found in input files.")
    exit()

combined_df = pd.concat(all_data, ignore_index=True)
print(f"✅ Combined total rows: {len(combined_df)}")

# Merge with reference on To Office Name
if 'to office name' not in combined_df.columns:
    print("❌ 'To Office Name' column not found in input files.")
    exit()

if 'office name' not in reference_df.columns or 'division' not in reference_df.columns:
    print("❌ 'Office Name' or 'Division' column missing in reference file.")
    exit()

merged_df = combined_df.merge(
    reference_df,
    how='left',
    left_on='to office name',
    right_on='office name'
)

# Remove unmatched rows (no division)
filtered_df = merged_df[merged_df['division'].notna()].copy()
print(f"✅ Matched rows after removing unmatched offices: {len(filtered_df)}")

# Prepare final output columns
final_columns = ['division', 'to office name', 'bag number', 'article count', 'bag type']
for col in final_columns:
    if col not in filtered_df.columns:
        filtered_df[col] = ""

filtered_df = filtered_df[final_columns]

# Sort by Division
filtered_df = filtered_df.sort_values(by='division')

# Split into PL and SP sheets
pl_df = filtered_df[filtered_df['bag type'].str.upper() == 'PL']
sp_df = filtered_df[filtered_df['bag type'].str.upper() == 'SP']

# Create summary
summary_df = (
    filtered_df.groupby(['division', 'bag type'])['article count']
    .sum()
    .reset_index()
    .pivot(index='division', columns='bag type', values='article count')
    .fillna(0)
    .reset_index()
)

# Save output
output_filename = "division_mapped_output.xlsx"
with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
    pl_df.to_excel(writer, index=False, sheet_name='PL Bags')
    sp_df.to_excel(writer, index=False, sheet_name='SP Bags')
    summary_df.to_excel(writer, index=False, sheet_name='Summary')

print(f"✅ Output Excel created with 3 sheets: {output_filename}")
print("   • PL Bags\n   • SP Bags\n   • Summary (Division totals)")

print("\n🎯 Process Completed Successfully!")
