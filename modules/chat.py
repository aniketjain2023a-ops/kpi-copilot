

from chatbot import ask_kpi_copilot

print("=" * 50)
print("KPI COPILOT CHAT")
print("Type 'exit' to quit")
print("=" * 50)

while True:
    question = input("\n> ")

    if question.lower() in ["exit", "quit"]:
        break

    try:
        answer = ask_kpi_copilot(question)

        print("\nKPI Copilot:")
        print(answer)

    except Exception as e:
        print(f"\nError: {e}")