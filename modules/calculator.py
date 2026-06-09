print("USING CALCULATOR FILE:", __file__)
import pandas as pd


def generate_kpi_analysis(settings_df, data_df):
    print("generate_kpi_analysis() CALLED")
    results = []

    for idx, setting_row in settings_df.iterrows():
        parameter = str(setting_row["Parameter Name"]).strip()
        annual_goal = setting_row["Annual Goal"]

        # Match rows by position instead of Parameter Name.
        # The Data Entry sheet uses formulas (=Settings!A2 etc.),
        # which pandas reads as NaN.
        if idx >= len(data_df):
            continue

        data_row = data_df.iloc[[idx]]

        monthly_data = data_row.iloc[0, 1:]
        print(f"[CALC] KPI: {parameter}")
        print(f"[CALC] Monthly Data:")
        print(monthly_data)
        print("-" * 50)



        valid_months = monthly_data.dropna()

        if len(valid_months) > 0:
            achieved = float(valid_months.iloc[-1])
        else:
            achieved = 0

        months_completed = monthly_data.count()
        months_remaining = 12 - months_completed

        remaining_target = annual_goal - achieved

        required_monthly_target = (
            remaining_target / months_remaining
            if months_remaining > 0
            else 0
        )

        parameter_lower = parameter.lower()

        # Percentage KPIs already represent percentages.
        if '%' in parameter:
            achievement_percent = achieved

        # Lower-is-better KPIs.
        elif any(x in parameter_lower for x in [
            'complaint',
            'energy consumption'
        ]):
            achievement_percent = (
                annual_goal / achieved * 100
                if achieved > 0 else 0
            )

            # Cap lower-is-better KPIs at 100%
            achievement_percent = min(achievement_percent, 100)

        # Standard higher-is-better KPIs.
        else:
            achievement_percent = (
                achieved / annual_goal * 100
                if annual_goal > 0 else 0
            )

        expected_percent = (
            months_completed / 12 * 100
        )

        variance = (
            achievement_percent - expected_percent
        )

        achievement_percent = max(
            0,
            min(round(achievement_percent, 2), 100)
        )

        if variance >= 0:
            status = "On Track"
        elif variance >= -10:
            status = "Monitor"
        else:
            status = "Critical"

        results.append({
            "Parameter Name": parameter,
            "Annual Goal": annual_goal,
            "Achieved": round(achieved, 2),
            "Remaining Target": round(remaining_target, 2),
            "Months Completed": months_completed,
            "Months Remaining": months_remaining,
            "Required Monthly Target": round(required_monthly_target, 2),
            "Achievement %": round(achievement_percent, 2),
            "Status": status
        })

    return pd.DataFrame(results)