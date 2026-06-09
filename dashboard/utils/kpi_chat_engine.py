import os

try:
    import google.generativeai as genai
except Exception:
    genai = None


def answer_question(question, df):
    if df.empty:
        return 'No KPI data is currently available.'

    if genai is None:
        return 'Gemini SDK not installed. Run: pip install google-generativeai'

    api_key = os.getenv('GEMINI_API_KEY')

    if not api_key:
        return 'GEMINI_API_KEY environment variable not configured.'

    try:
        genai.configure(api_key=api_key)

        model = genai.GenerativeModel('gemini-2.5-flash')

        context = df.to_string(index=False)

        prompt = f'''
You are KPI Copilot, an AI management analyst.

KPI Data:
{context}

User Question:
{question}

Provide:
- Direct answer
- Key insights
- Risks
- Recommended actions
'''

        response = model.generate_content(prompt)

        return response.text

    except Exception as e:
        return f'Gemini error: {e}'
