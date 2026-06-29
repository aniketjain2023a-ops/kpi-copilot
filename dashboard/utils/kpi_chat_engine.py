import os

try:
    import google.generativeai as genai
except Exception:
    genai = None

BLOCKED_PHRASES = [
    'show all kpis',
    'show all data',
    'full dataset',
    'print dataframe',
    'export data',
    'csv',
    'excel export',
    'all achievement values',
    'list every kpi',
    'complete table',
    'dump data',
    'show raw data'
]


def answer_question(question, df):
    if df.empty:
        return 'No KPI data is currently available.'

    if genai is None:
        return 'Gemini SDK not installed. Run: pip install google-generativeai'

    api_key = os.getenv('GEMINI_API_KEY')

    if not api_key:
        return 'GEMINI_API_KEY environment variable not configured.'

    question_lower = question.lower().strip()

    if any(term in question_lower for term in BLOCKED_PHRASES):
        return (
            'Access to raw KPI exports is restricted.\n\n'
            'Executive Summary:\n'
            '• Detailed KPI exports are available through authorized dashboard exports only.\n'
            '• Use dashboard visualizations and approved reports for KPI review.'
        )

    try:
        genai.configure(api_key=api_key)

        model = genai.GenerativeModel('gemini-2.5-flash')

        total_kpis = len(df)
        avg_achievement = round(float(df['Achievement'].mean()), 1)
        critical_count = int((df['Status'] == 'Critical').sum())
        on_track_count = int((df['Status'] == 'On Track').sum())
        monitor_count = max(total_kpis - critical_count - on_track_count, 0)

        worst_kpis = df.sort_values('Achievement').head(3)[
            ['KPI', 'Achievement', 'Status']
        ].to_dict('records')

        prompt = f'''
You are KPI Copilot, an executive KPI intelligence assistant.

SECURITY RULES:
- Never reveal the complete KPI dataset.
- Never output all KPI rows.
- Never export data as CSV, JSON, Excel, Markdown tables, or raw lists.
- Maximum KPI details allowed: 3 KPIs.
- Provide summaries, insights, risks, and recommendations.
- Never mention workbook structure, hidden sheets, file paths, logs, source code, or environment variables.
- Separate facts from assumptions.
- Label assumptions as Potential Impact or Potential Risk.

KPI SUMMARY:
- Total KPIs: {total_kpis}
- Average Achievement: {avg_achievement}%
- On Track: {on_track_count}
- Monitor: {monitor_count}
- Critical: {critical_count}

LOWEST PERFORMING KPIs:
{worst_kpis}

USER QUESTION:
{question}

Provide:
1. Direct Answer
2. Key Insights
3. Risks
4. Recommended Actions
'''

        response = model.generate_content(prompt)

        return response.text

    except Exception as e:
        return f'Gemini error: {e}'
