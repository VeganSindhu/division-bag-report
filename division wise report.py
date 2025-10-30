import streamlit as st
import pandas as pd
from pathlib import Path
import io  # For in-memory download

st.title("🔄 Division-wise Report Generator")
st.write("Processes files from repo: Loads reference, scans for inputs (CSV/XLS/XLSX), generates output.")

# Your original reference loading (direct from repo)
reference_filename = "division wis.xlsx"
reference_path = Path(reference_filename)
if not reference_path.exists():
    st.error(f"❌ Reference file '{reference_filename}' not found in repo. Add it and redeploy.")
    st.stop()
reference_df = pd.read_excel(reference_path)
reference_df.columns = reference_df.columns.str.strip().str.lower()
st.success(f"✅ Loaded reference: {reference_filename} ({len(reference_df)} rows)")

# Your original input scanning (from repo)
input_files = [
    f for f in Path(".").glob("*.*")
    if f.suffix.lower() in [".csv", ".xls", ".xlsx"] and f.name != reference_filename
]
if not input_files:
    st.warning("⚠️ No input files (CSV/XLS/XLSX) found in repo. Add them and redeploy.")
    st.stop()
st.info(f"📂 Found {len(input_files)} input files: {', '.join(f.name for f in input_files)}")

# Your original combining logic
all_data = []
for file in input_files:
    try:
        if file.suffix.lower() == ".csv":
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
        df.columns = df.columns.str.strip().str.lower()
        fname = file.name.lower()
        if "set1" in fname:
            df["bag type"] = "PL"
        elif "set2" in fname:
            df["bag type"] = "SP"
        all_data.append(df)
    except Exception as e:
        st.error(f"❌ Error reading {file.name}: {e}")
        st.stop()
if not all_data:
    st.warning("⚠️ No valid data in inputs.")
    st.stop()
combined_df = pd.concat(all_data, ignore_index=True)
st.success(f"✅ Combined {len(combined_df)} rows")

# Your original merge/filter
if 'to office name' not in combined_df.columns:
    st.error("❌ 'To Office Name' column not found in inputs.")
    st.stop()
if 'office name' not in reference_df.columns or 'division' not in reference_df.columns:
    st.error("❌ 'Office Name' or 'Division' missing in reference.")
    st.stop()
merged_df = combined_df.merge(
    reference_df, how='left', left_on='to office name', right_on='office name'
)
filtered_df = merged_df[merged_df['division'].notna()].copy()
st.info(f"✅ Matched: {len(filtered_df)} rows")

# Your original output prep
final_columns = ['division', 'to office name', 'bag number', 'article count', 'bag type']
for col in final_columns:
    if col not in filtered_df.columns:
        filtered_df[col] = ""
filtered_df = filtered_df[final_columns].sort_values(by='division')

# Split and summary (your code)
pl_df = filtered_df[filtered_df['bag type'].str.upper() == 'PL']
sp_df = filtered_df[filtered_df['bag type'].str.upper() == 'SP']
summary_df = (
    filtered_df.groupby(['division', 'bag type'])['article count']
    .sum().reset_index()
    .pivot(index='division', columns='bag type', values='article count')
    .fillna(0).reset_index()
)

# Run button (processes on click)
if st.button("🚀 Generate Report", type="primary"):
    with st.spinner("Processing..."):
        # Create in-memory Excel (your output logic)
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            pl_df.to_excel(writer, index=False, sheet_name='PL Bags')
            sp_df.to_excel(writer, index=False, sheet_name='SP Bags')
            summary_df.to_excel(writer, index=False, sheet_name='Summary')
        output_buffer.seek(0)

        # Download
        st.download_button(
            label="📥 Download: division_mapped_output.xlsx",
            data=output_buffer.getvalue(),
            file_name="division_mapped_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Preview summary
        st.subheader("📊 Preview: Division Summary")
        st.dataframe(summary_df, use_container_width=True)

        st.success("🎯 Report generated! Download above. Add more input files to repo for updates.")
else:
    st.info("👆 Click 'Generate Report' to process files from repo.")
