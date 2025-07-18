import streamlit as st
import pandas as pd
import numpy as np
import io

st.title("Schedule RESP Analyzer")

uploaded_file = st.file_uploader(
    "➡️ Copy and paste the schedule from **P6**, and make sure the schedule includes the following columns:\n"
    "- `Original Duration`\n"
    "- `Actual Duration`\n"
    "- `Activity Status`\n"
    "- `Resp` (this one could be project specific like `Resp6`) \n"
    "\n 📂 **Upload Excel file(s) (.xlsx)** ",
    type="xlsx",
    accept_multiple_files=False
)

min_activities = st.number_input(
    "Minimum # of Activities per RESP Group (to include in summary)", 
    min_value=1, 
    value=50, 
    step=1
)


def process_file(file):
    df = pd.read_excel(file)
    df.columns = [col.strip() for col in df.columns]

    # Clean + rename columns
    if 'G - Resp' in df.columns and df['G - Resp'].isna().all():
        df = df.drop(columns=['G - Resp'])

    for col in df.columns:
        if 'Original Duration' in col:
            df = df.rename(columns={col: 'OD'})
        elif 'Actual Duration' in col:
            df = df.rename(columns={col: 'ACDur'})
        elif 'G - Resp' not in col and 'resp' in col.lower():
            df = df.rename(columns={col: 'G - Resp'})

    if 'G - Resp' not in df.columns or df['G - Resp'].isna().all():
        return None  # skip bad files

    df = df.dropna(subset=['OD', 'ACDur'])
    df = df[df['Activity Status'] == 'Completed']
    df = df[df['OD'] != 0]
    df['ACDur'] = df.apply(lambda x: min(x['ACDur'], 2 * x['OD']), axis=1)
    df = df.dropna(subset=['G - Resp'])

    results = []
    for resp in df['G - Resp'].unique():
        filtered = df[df['G - Resp'] == resp]
        if len(filtered) < min_activities:
            continue

        df_min = filtered[filtered['ACDur'] < filtered['OD']]
        df_max = filtered[filtered['ACDur'] > filtered['OD']]

        min_ratio = df_min['ACDur'].sum() / df_min['OD'].sum() if not df_min.empty else np.nan
        ml_ratio = filtered['ACDur'].sum() / filtered['OD'].sum()
        max_ratio = df_max['ACDur'].sum() / df_max['OD'].sum() if not df_max.empty else np.nan

        results.append({
            'G - Resp': resp,
            'Min': round(min_ratio, 4) if pd.notna(min_ratio) else 1,
            'Most Likely': round(ml_ratio, 4) if pd.notna(ml_ratio) else None,
            'Max': round(max_ratio, 4) if pd.notna(max_ratio) else 2
        })

    return pd.DataFrame(results)

if uploaded_file:
    result_df = process_file(uploaded_file)
    
    if result_df is not None:
        st.write("✅ Processed summary:")
        st.dataframe(result_df)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            result_df.to_excel(writer, index=False)
        st.download_button("Download Summary", output.getvalue(), file_name="Resp-Ratios.xlsx")
    else:
        st.warning("No valid RESP data found in uploaded file.")


st.header("How the Ratios are Calculated")

st.markdown("""
- **Min**: Sum of `Actual Duration` / `Original Duration` for all activities that finished **early**  
- **Most Likely**: Sum of `Actual Duration` / `Original Duration` for **all completed activities**  
- **Max**: Sum of `Actual Duration` / `Original Duration` for all activities that finished **late**
""")
