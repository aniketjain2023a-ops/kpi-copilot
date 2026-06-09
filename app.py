import pandas as pd

from modules.calculator import generate_kpi_analysis
from modules.writer import write_analysis_sheet


def main():
    print("Starting KPI Copilot...")
    workbook_path = "reports/kpi_target_tracker.xlsx"

    settings_df = pd.read_excel(
        workbook_path,
        sheet_name="Settings"
    )

    data_df = pd.read_excel(
        workbook_path,
        sheet_name="Data Entry"
    )

    try:
        analysis_df = generate_kpi_analysis(
            settings_df,
            data_df
        )

        if analysis_df.empty:
            print("ERROR: No KPI analysis generated.")
            return

        print("✓ KPI calculations completed")

        write_analysis_sheet(
            workbook_path,
            analysis_df
        )

        print("✓ Calculations sheet updated")
        print("\nKPI Copilot completed successfully.")

    except Exception as e:
        print(f"ERROR: {e}")
        raise


if __name__ == "__main__":
    main()