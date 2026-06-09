

import streamlit as st
import pandas as pd
from pathlib import Path

from utils.kpi_chat_engine import answer_question

st.title('🤖 KPI Copilot Chatbot')
st.caption('Powered by Gemini AI for KPI analysis, insights, risks, forecasting, and recommendations.')

project_root = Path(__file__).resolve().parents[2]
workbook_path = project_root / 'reports' / 'kpi_target_tracker.xlsx'

try:
    calc_df = pd.read_excel(
        workbook_path,
        sheet_name='Calculations',
        header=None
    )

    header_row = None

    for idx in range(min(20, len(calc_df))):
        row_values = [str(v).lower() for v in calc_df.iloc[idx].tolist()]

        if any('parameter' in v for v in row_values):
            header_row = idx
            break

    calc_df.columns = calc_df.iloc[header_row]
    calc_df = calc_df.iloc[header_row + 1:].reset_index(drop=True)

    cols = [str(c) for c in calc_df.columns]

    kpi_col = next((c for c in cols if 'parameter' in c.lower()), cols[0])
    achievement_col = next((c for c in cols if 'achievement' in c.lower()), cols[1])
    status_col = next((c for c in cols if 'status' in c.lower()), cols[-1])

    df = pd.DataFrame({
        'KPI': calc_df[kpi_col].astype(str),
        'Achievement': pd.to_numeric(calc_df[achievement_col], errors='coerce'),
        'Status': calc_df[status_col].astype(str)
    })

    df = df.dropna(subset=['Achievement']).reset_index(drop=True)

except Exception as e:
    st.error(f'Unable to load KPI data: {e}')
    df = pd.DataFrame(columns=['KPI', 'Achievement', 'Status'])

st.markdown('### Suggested Questions')

suggestions = [
    'Give me a management summary',
    'Which KPI is performing worst?',
    'Which KPI is performing best?',
    'Show KPIs below 75%',
    'What are the key risks?',
    'What actions should management take?'
]

cols = st.columns(3)
for i, question in enumerate(suggestions):
    if cols[i % 3].button(question):
        st.session_state['selected_question'] = question

if 'chat_history' not in st.session_state:
    st.session_state.chat_history = []

prompt = st.chat_input(
    'Ask anything about KPI performance, targets, trends, risks, forecasts, or recommendations...'
)

if st.session_state.get('selected_question'):
    prompt = st.session_state.pop('selected_question')

if prompt:
    try:
        answer = answer_question(prompt, df)
    except Exception as e:
        answer = f'Chatbot error: {e}'

    st.session_state.chat_history.append((prompt, answer))

for question, answer in st.session_state.chat_history:
    with st.chat_message('user'):
        st.write(question)

    with st.chat_message('assistant'):
        st.write(answer)