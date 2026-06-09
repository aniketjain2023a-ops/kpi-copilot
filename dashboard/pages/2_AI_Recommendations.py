

import streamlit as st
import pandas as pd
from pathlib import Path

from utils.kpi_chat_engine import answer_question

st.title('🤖 AI Recommendations')
st.caption('Gemini-powered KPI insights and management recommendations')

project_root = Path(__file__).resolve().parents[2]
workbook_path = project_root / 'reports' / 'kpi_target_tracker.xlsx'

try:
    calc_df = pd.read_excel(workbook_path, sheet_name='Calculations', header=None)

    kpi_df = calc_df.iloc[4:14, [0, 7, 8]].copy()
    kpi_df.columns = ['KPI', 'Achievement', 'Status']

    kpi_df['Achievement'] = pd.to_numeric(
        kpi_df['Achievement'], errors='coerce'
    )

    critical_df = kpi_df[
        kpi_df['Status'].astype(str).isin(['Critical', 'Monitor'])
    ].sort_values('Achievement')

    st.subheader('Management Focus Areas')

    if not critical_df.empty:
        st.dataframe(critical_df, width='stretch')
    else:
        st.success('All KPIs are currently on track.')

    st.divider()

    col1, col2 = st.columns(2)

    with col1:
        if st.button('📊 Generate Management Summary'):
            st.session_state['ai_output'] = answer_question(
                'Give me a management summary',
                kpi_df
            )

        if st.button('⚠️ Generate Risk Report'):
            st.session_state['ai_output'] = answer_question(
                'Which KPIs are at risk and what actions should management take?',
                kpi_df
            )

    with col2:
        if st.button('👔 Generate Executive Summary'):
            st.session_state['ai_output'] = answer_question(
                'Give me an executive summary for leadership',
                kpi_df
            )

        if st.button('🚀 Generate Action Plan'):
            st.session_state['ai_output'] = answer_question(
                'Create an action plan to improve low performing KPIs',
                kpi_df
            )

    if 'ai_output' in st.session_state:
        st.divider()
        st.subheader('AI Analysis')
        st.write(st.session_state['ai_output'])

except Exception as e:
    st.error(f'Unable to load KPI data: {e}')