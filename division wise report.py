import streamlit as st
import pandas as pd
from pathlib import Path
import io  # For in-memory download

st.title("🔄 Division-wise Report Generator")
st.write("Loads reference from repo, upload inputs (CSV/XLS/XLSX) to generate report.")

# Load reference from repo (unchanged)
reference_filename = "division wis.xlsx"
reference_path = Path(reference_filename)
if not reference_path.exists():
    st.error(f"❌ Reference file '{reference_filename}' not found in repo. Add it and redeploy.")
    st.stop()
reference_df = pd.read_excel(reference_path)
reference_df.columns = reference_df.columns.str.strip().str.lower()
st.success(f"✅ Loaded reference: {reference_filename} ({len(reference_df)} rows)")

# File upload for inputs (new: multi-uploader)
input_files = st.file_uploader(
    "Upload Input Files (CSV/XLS/XLSX—up to 4 or more)",
    accept_multiple_files=True,
    type=['csv', 'xls', 'xlsx'],
    help="Upload your 4 Excel files here (e.g., set1_data.xlsx, set2_data.xlsx, etc.). Bag types auto-detected from filenames."
)

if not input_files:
    st.warning("👆 Upload at least one input file to proceed.")
    st.stop()

st.info(f"📂 Uploaded {len(input_files)} input files: {', '.join(f.name for f in input_files)}")

# Combine inputs (adapted from your original)
all_data = []
for file in input_files:
    try:
        if file.name.lower().endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
        df.columns = df.columns.str.strip().str.lower()
        fname = file.name.lower()
        if "set1" in fname:
            df["bag type"] = "PL"
        elif "set2" in fname:
            df["bag type"] = "SP"
        else:
            df["bag type"] = "Unknown"  # Fallback—edit if needed
        all_data.append(df)
    except Exception as e:
        st.error(f"❌ Error reading {file.name}: {e}")
        st.stop()
if not all_data:
    st.warning("⚠️ No valid data in uploads.")
    st.stop()
combined_df = pd.concat(all_data, ignore_index=True)
st.success(f"✅ Combined {len(combined_df)} rows from uploads")

# Merge/filter (your original)
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

# Output prep (your original)
final_columns = ['division', 'to office name', 'bag number', 'article count', 'bag type']
for col in final_columns:
    if col not in filtered_df.columns:
        filtered_df[col] = ""
filtered_df = filtered_df[final_columns].sort_values(by='division')

# Split and summary (your original)
pl_df = filtered_df[filtered_df['bag type'].str.upper() == 'PL']
sp_df = filtered_df[filtered_df['bag type'].str.upper() == 'SP']
summary_df = (
    filtered_df.groupby(['division', 'bag type'])['article count']
    .sum().reset_index()
    .pivot(index='division', columns='bag type', values='article count')
    .fillna(0).reset_index()
)

# Generate button (processes on click)
if st.button("🚀 Generate Report", type="primary"):
    with st.spinner("🔄 Processing..."):
        # In-memory Excel (your output)
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

        # Preview
        st.subheader("📊 Preview: Division Summary")
        st.dataframe(summary_df, use_container_width=True)

        st.success("🎯 Report ready! Download above.")
else:
    st.info("👆 Upload files, then click 'Generate Report'.")
