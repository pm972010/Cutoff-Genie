import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO

st.title("📊 CUTOFF GENIE")

# 1. File uploader
uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=['xlsx'])

# 2. Date input
start_date = st.date_input("Select Start Date", format="DD-MM-YYYY")
end_date = st.date_input("Select End Date", format="DD-MM-YYYY")

# 3. Execute button
if uploaded_file and st.button("🚀 Execute"):
    try:
        df = pd.read_excel(uploaded_file, skiprows=7)
        df = df.iloc[:, 1:12]
        df.iloc[:, 0] = df.iloc[:, 0].ffill()
        df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0].astype(str) + ' ' + df.iloc[:, 4].astype(str))
        df.drop(df.columns[4], axis=1, inplace=True)
        max_data_date = df.iloc[:, 0].max().date()
        end_date = min(end_date, max_data_date)

        # Step 3: Filter valid dates — only process dates where data exists
        valid_dates = []
        for single_date in [start_date + timedelta(days=x) for x in range((end_date - start_date).days + 1)]:
            target_datetime = datetime.combine(single_date, datetime.min.time())

            # Only add if any data exists on or before that datetime
            if not df[df.iloc[:, 0] <= target_datetime].empty:
                valid_dates.append(single_date)

        # Step 4: Prepare results
        results = []
        group_col = df.columns[1]

        # Step 5: Iterate through valid dates and groups
        for single_date in valid_dates:
            target_datetime = datetime.combine(single_date.date(), datetime.min.time())

            for group_val, group_df in df.groupby(group_col):
                group_df = group_df.sort_values(by=df.columns[0])
                subset = group_df[group_df.iloc[:, 0] <= target_datetime]
                if not subset.empty:
                    row = subset.iloc[-1]
                    if row.iloc[4]=="CL":
                        values = row.iloc[5:10].copy()
                    else:
                        values = row.iloc[5:10].copy()
                        values.iloc[0] = 'ATG'

                    # Append to result: [Date, Group, Condition, Col6–11]
                    results.append([
                        target_datetime.date(),
                        row.iloc[1],
                        row.iloc[4],
                        *values.values
                    ])

        # Step 6: Create final DataFrame
        final_df = pd.DataFrame(results, columns=[
            'Date',
            df.columns[1],        # Group column
            df.columns[4],        # CL/OP
            *df.columns[5:10]     # Columns 6 to 11
        ])

        # Show date only once per block
        final_df['Date'] = final_df['Date'].astype(str)
        final_df['Date'] = final_df['Date'].mask(final_df['Date'].duplicated(), '')

        # Step 7: Display DataFrame
        st.subheader("📄 Processed Data")
        st.dataframe(final_df, use_container_width=True)

        # Step 8: Download as Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Result')
        st.download_button(
            label="⬇️ Download Excel",
            data=output.getvalue(),
            file_name="grouped_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error processing the file: {e}")
else:
    st.info("📂 Please upload an Excel file to begin.")
    
    

