

import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

st.title('📈 Historical Analysis')
st.caption('Analyze KPI trends, performance history, and future trajectory')

project_root = Path(__file__).resolve().parents[2]
workbook_path = project_root / 'reports' / 'kpi_target_tracker.xlsx'

try:
    data_df = pd.read_excel(workbook_path, sheet_name='Data Entry', header=None)

    # Data Entry sheet structure detected from workbook
    month_headers = data_df.iloc[0].tolist()

    known_kpis = [
        'Production Volume',
        'Sales Revenue',
        'Dealer Addition',
        'Customer Complaints',
        'Safety Observations',
        'Energy Consumption',
        'Inventory Turnover',
        'OTIF Delivery (%)',
        'Quality Score (%)',
        'Cost Reduction ($)'
    ]

    # First 10 rows after header correspond to KPI values
    kpi_rows = data_df.iloc[1:11].reset_index(drop=True)

    kpi_df = pd.DataFrame()
    kpi_df['KPI'] = known_kpis

    for col_idx in range(1, min(len(month_headers), len(kpi_rows.columns))):
        month_name = str(month_headers[col_idx])
        kpi_df[month_name] = kpi_rows.iloc[:, col_idx].values

    selected_kpi = st.selectbox(
        'Select KPI',
        kpi_df['KPI'].tolist()
    )

    selected_row = kpi_df[
        kpi_df['KPI'] == selected_kpi
    ].iloc[0]

    value_cols = [c for c in kpi_df.columns if c != 'KPI']

    trend_df = pd.DataFrame({
        'Month': value_cols,
        'Value': [selected_row[col] for col in value_cols]
    })

    trend_df['Value'] = pd.to_numeric(trend_df['Value'], errors='coerce')
    trend_df = trend_df.dropna()

    if trend_df.empty:
        st.warning('No historical monthly data found for this KPI.')
        st.stop()

    col1, col2 = st.columns(2)

    with col1:
        st.metric('Latest Value', round(float(trend_df['Value'].iloc[-1]), 2))

    with col2:
        growth = 0
        if len(trend_df) > 1 and trend_df['Value'].iloc[0] != 0:
            growth = (
                (trend_df['Value'].iloc[-1] - trend_df['Value'].iloc[0])
                / trend_df['Value'].iloc[0]
            ) * 100

        st.metric('Growth %', f'{growth:.2f}%')

    st.subheader('KPI Trend')

    fig = px.line(
        trend_df,
        x='Month',
        y='Value',
        markers=True,
        title=f'{selected_kpi} Trend'
    )

    st.plotly_chart(fig, width='stretch')

    st.subheader('Forecast')

    if len(trend_df) >= 2:
        last_value = trend_df['Value'].iloc[-1]
        change = trend_df['Value'].diff().mean()

        forecast_df = pd.DataFrame({
            'Month': ['Forecast +1', 'Forecast +2', 'Forecast +3'],
            'Value': [
                last_value + change,
                last_value + (change * 2),
                last_value + (change * 3)
            ]
        })

        forecast_chart = pd.concat([
            trend_df[['Month', 'Value']],
            forecast_df
        ])

        fig2 = px.line(
            forecast_chart,
            x='Month',
            y='Value',
            markers=True,
            title='3-Month Forecast'
        )

        st.plotly_chart(fig2, width='stretch')
        st.dataframe(forecast_df, width='stretch')

except Exception as e:
    st.error(f'Unable to load historical analysis: {e}')