import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

st.title('📊 KPI Dashboard')

# Set up paths
project_root = Path(__file__).resolve().parents[2]
workbook_path = project_root / 'reports' / 'kpi_target_tracker.xlsx'

# Try loading the Calculations sheet
try:
    calc_df = pd.read_excel(workbook_path, sheet_name='Calculations', header=None)
    st.write('DEBUG - First 20 rows from Calculations sheet')
    st.dataframe(calc_df.head(20), width='stretch')
    # Find the header row containing 'Parameter'
    header_row = None
    for idx, row in calc_df.iterrows():
        if row.astype(str).str.contains('Parameter', case=False).any():
            header_row = idx
            break
    if header_row is not None:
        # Assign column names
        columns = calc_df.iloc[header_row].tolist()
        # Build a dataframe with the correct columns
        data_start = header_row + 1
        kpi_df = calc_df.iloc[data_start:].copy()
        kpi_df.columns = columns
        # The Calculations sheet structure is fixed:
        # Column 0 = KPI Name
        # Column 7 = Achievement %
        # Column 8 = Status

        kpi_df = pd.DataFrame({
            'KPI': kpi_df.iloc[:, 0],
            'Achievement': kpi_df.iloc[:, 7],
            'Status': kpi_df.iloc[:, 8]
        })

        if 'KPI' not in kpi_df.columns:
            kpi_df['KPI'] = kpi_df.iloc[:, 0]

        if 'Achievement' not in kpi_df.columns:
            achievement_candidates = [
                c for c in kpi_df.columns
                if any(x in str(c).lower() for x in ['achievement', '%', 'actual', 'score'])
            ]
            if achievement_candidates:
                kpi_df['Achievement'] = kpi_df[achievement_candidates[0]]
            else:
                kpi_df['Achievement'] = None

        if 'Status' not in kpi_df.columns:
            kpi_df['Status'] = 'Unknown'

        # Remove duplicate columns that can be created by rename operations
        kpi_df = kpi_df.loc[:, ~kpi_df.columns.duplicated()].copy()
        # Drop rows where KPI is missing
        kpi_df = kpi_df[kpi_df['KPI'].notna()]
        # Clean up Achievement column
        if isinstance(kpi_df['Achievement'], pd.DataFrame):
            kpi_df['Achievement'] = kpi_df['Achievement'].iloc[:, 0]

        kpi_df['Achievement'] = pd.to_numeric(kpi_df['Achievement'], errors='coerce')
        kpi_df = kpi_df.reset_index(drop=True)

        # Remove summary and narrative rows
        kpi_df = kpi_df[
            ~kpi_df['KPI'].astype(str).str.contains(
                'SUMMARY|Generated|Total KPI|On Track|Critical|Monitor|Narrative',
                case=False,
                na=False
            )
        ]

        kpi_df = kpi_df.dropna(subset=['Achievement'])
    else:
        kpi_df = pd.DataFrame(columns=['KPI', 'Achievement', 'Status'])
except Exception as e:
    st.error(f'Unable to load KPI workbook: {e}')
    st.exception(e)
    kpi_df = pd.DataFrame(columns=['KPI', 'Achievement', 'Status'])

# Show metrics if data exists
if not kpi_df.empty:
    total_kpis = len(kpi_df)
    avg_achievement = kpi_df['Achievement'].mean()
    on_track = (kpi_df['Status'].astype(str).str.lower() == 'on track').sum()
    needs_attention = (
        kpi_df['Status'].astype(str).str.lower().isin(['monitor', 'critical'])
    ).sum()
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total KPIs", total_kpis)
    col2.metric("Average Achievement", f"{avg_achievement:.1f}%" if pd.notnull(avg_achievement) else "N/A")
    col3.metric("On Track", on_track)
    col4.metric("Needs Attention", needs_attention)

    st.markdown("---")

    left, right = st.columns(2)
    with left:
        fig = px.bar(
            kpi_df,
            x='KPI',
            y='Achievement',
            color='Status',
            title='KPI Performance',
            labels={'Achievement': 'Achievement (%)'},
            text='Achievement',
        )
        fig.update_layout(xaxis_title='', yaxis_title='Achievement (%)', showlegend=True)
        left.plotly_chart(fig, use_container_width=True)
    with right:
        st.subheader("Current Status")
        st.dataframe(kpi_df, use_container_width=True)

    st.markdown("---")
    st.subheader("Status Distribution")
    status_counts = kpi_df['Status'].value_counts(dropna=False)
    pie_fig = px.pie(
        names=status_counts.index.astype(str),
        values=status_counts.values,
        title='KPI Status Distribution',
        color_discrete_sequence=px.colors.qualitative.Set2
    )
    st.plotly_chart(pie_fig, use_container_width=True)
else:
    st.write('Detected columns:', list(kpi_df.columns))
    if not kpi_df.empty:
        st.dataframe(kpi_df.head())
    st.info("No KPI data available to display.")

st.markdown("---")
with open(workbook_path, 'rb') as f:
    st.download_button(
        '📥 Download Updated Excel',
        f.read(),
        file_name='kpi_target_tracker.xlsx'
    )