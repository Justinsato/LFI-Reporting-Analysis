import pandas as pd
import joblib
import os
from datetime import datetime
import logging

def main(req):
    try:
        file = req.files["file"]
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()
        today = datetime.today()

        df["Move-in"] = pd.to_datetime(df["Move-in"], errors="coerce")
        df["Tenure (Years)"] = ((today - df["Move-in"]).dt.days / 365).fillna(0)

        pressure_df = pd.read_csv("rent_pressure.csv")
        pressure_series = pressure_df["Rent Pressure (%)"].tolist()

        if "Move-out" in df.columns:
            df["Move-out"] = pd.to_datetime(df["Move-out"], errors="coerce")

        def turnover_prob(tenure):
            penalty = sum(pressure_series[:int(tenure)]) / 100 if tenure >= 1 else 0
            return max(0.05, 0.5 - penalty)

        df["Turnover Probability"] = df.apply(
            lambda row: turnover_prob(row["Tenure (Years)"])
            if (row.get("Status") != "Vacant" and pd.isna(row.get("Move-out")))
            else row.get("Turnover Probability", 0.5), axis=1)

        from io import BytesIO
        output = BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        return {
            "status": 200,
            "headers": { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
            "body": output.read()
        }

    except Exception as e:
        logging.exception("Processing failed")
        return {
            "status": 500,
            "body": f"Error: {str(e)}"
        }
