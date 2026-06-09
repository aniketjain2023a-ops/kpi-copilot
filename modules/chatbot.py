import pandas as pd
import ollama
from pathlib import Path


def ask_kpi_copilot(question):
    project_root = Path(__file__).resolve().parent.parent
    workbook_path = project_root / "reports" / "kpi_target_tracker.xlsx"

    df = pd.read_excel(
        workbook_path,
        sheet_name="Calculations"
    )
    df.columns = [str(c).strip() for c in df.columns]

    q = question.lower()

    # Instant rule-based answers
    if "highest risk" in q or "risk kpi" in q:
        if "Status" in df.columns:
            critical = df[df["Status"].astype(str).str.contains("Critical", case=False, na=False)]
            if not critical.empty:
                return f"Highest risk KPI: {critical.iloc[0]['Parameter Name']}"

    if "maximum effort" in q or "highest effort" in q:
        if "Required Monthly Target" in df.columns:
            row = df.loc[df["Required Monthly Target"].idxmax()]
            return f"Highest effort KPI: {row['Parameter Name']}"

    if "top performer" in q or "best kpi" in q:
        if "Achievement %" in df.columns:
            row = df.loc[df["Achievement %"].idxmax()]
            return f"Top performer: {row['Parameter Name']} ({row['Achievement %']}%)"

    if "declining" in q and "Trend" in df.columns:
        count = len(df[df["Trend"].astype(str).str.contains("Declining", na=False)])
        return f"Declining KPIs: {count}"

    useful_cols = [
        col for col in [
            "Parameter Name",
            "Achievement %",
            "Required Monthly Target",
            "Status",
            "Trend"
        ]
        if col in df.columns
    ]

    if useful_cols:
        context = df[useful_cols].head(15).to_dict("records")
    else:
        context = df.head(15).to_dict("records")

    prompt = f"""You are KPI Copilot.
Answer in 1-3 short sentences.
Use only the KPI data below.

DATA: {context}

QUESTION: {question}
"""

    response = ollama.chat(
        model="llama3.2",
        messages=[
            {
                "role": "user",
                "content": prompt
            }
        ],
        options={
            "temperature": 0,
            "num_predict": 80,
            "num_ctx": 1024,
        }
    )

    return response["message"]["content"]