from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
RESULT_FILE = "output/Backtest_Result.xlsx"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs("output", exist_ok=True)


def categorize(cagr, thresholds):
    if cagr < thresholds['extreme_bearish']:
        return "Extreme Bearish"
    elif thresholds['extreme_bearish'] <= cagr < thresholds['bearish']:
        return "Bearish"
    elif thresholds['bearish'] <= cagr < thresholds['sideways_bearish']:
        return "Sideways Bearish"
    elif thresholds['sideways_bearish'] <= cagr < thresholds['neutral']:
        return "Neutral"
    elif thresholds['neutral'] <= cagr < thresholds['bullish']:
        return "Bullish"
    else:
        return "Extreme Bullish"


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        try:
            file = request.files["file"]
            holding_period = int(request.form.get("holding_period", 250))
            capital = float(request.form.get("capital", 2500000))

            thresholds = {
                "extreme_bearish": float(request.form.get("cutoff_extreme_bearish", 0)),
                "bearish": float(request.form.get("cutoff_bearish", 0.06)),
                "sideways_bearish": float(request.form.get("cutoff_sideways_bearish", 0.10)),
                "neutral": float(request.form.get("cutoff_neutral", 0.12)),
                "bullish": float(request.form.get("cutoff_bullish", 0.15))
            }

            units = {
                "extreme_bearish": float(request.form.get("units_extreme_bearish", 2)),
                "bearish": float(request.form.get("units_bearish", 1)),
                "sideways_bearish": float(request.form.get("units_sideways_bearish", 0.5)),
                "bullish": float(request.form.get("exit_units_bullish", 0.5)),
                "extreme_bullish": float(request.form.get("exit_units_extreme_bullish", 1))
            }

            if file:
                filepath = os.path.join(UPLOAD_FOLDER, file.filename)
                file.save(filepath)

                df = pd.read_excel(filepath, sheet_name="Sheet1", skiprows=2, engine="openpyxl")
                df.columns = df.columns.str.strip()
                print("COLUMNS READ >>>", df.columns.tolist())
                df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
                df = df.dropna(subset=["Date", "Close"]).reset_index(drop=True)

                cagr_records = []
                for i in range(len(df) - holding_period):
                    entry_row = df.iloc[i]
                    exit_row = df.iloc[i + holding_period]
                    entry_price = entry_row["Close"]
                    exit_price = exit_row["Close"]
                    cagr = (exit_price / entry_price) ** (252 / holding_period) - 1
                    cagr_records.append({
                        "Date": entry_row["Date"],
                        "Close": entry_price,
                        "Exit Date": exit_row["Date"],
                        "Exit Price": exit_price,
                        "Annualized CAGR": cagr,
                        "Category": categorize(cagr, thresholds)
                    })

                cagr_df = pd.DataFrame(cagr_records)

                total_units = 0
                invested_capital = 0
                total_withdrawals = 0
                cash = capital
                results = []

                for _, row in cagr_df.iterrows():
                    date = row["Date"]
                    close_price = row["Close"]
                    category = row["Category"]
                    buy = sell = 0

                    if category == "Extreme Bearish" and cash >= units['extreme_bearish'] * close_price:
                        buy = units['extreme_bearish']
                    elif category == "Bearish" and cash >= units['bearish'] * close_price:
                        buy = units['bearish']
                    elif category == "Sideways Bearish" and cash >= units['sideways_bearish'] * close_price:
                        buy = units['sideways_bearish']

                    invested_capital += buy * close_price
                    cash -= buy * close_price

                    if category == "Extreme Bullish" and total_units >= units['extreme_bullish']:
                        sell = units['extreme_bullish']
                    elif category == "Bullish" and total_units >= units['bullish']:
                        sell = units['bullish']

                    total_withdrawals += sell * close_price
                    cash += sell * close_price

                    total_units += buy - sell
                    portfolio_value = total_units * close_price

                    results.append([
                        date, category, close_price, buy, sell, total_units,
                        portfolio_value, invested_capital, total_withdrawals, cash
                    ])

                result_df = pd.DataFrame(results, columns=[
                    "Date", "Category", "Close Price", "Units Bought", "Units Sold", "Total Units Held",
                    "Portfolio Value", "Total Invested", "Total Withdrawn", "Remaining Cash"
                ])

                final_val = result_df.iloc[-1]["Portfolio Value"] + result_df.iloc[-1]["Remaining Cash"] + result_df.iloc[-1]["Total Withdrawn"]
                years = (result_df["Date"].max() - result_df["Date"].min()).days / 365.25
                cagr = (final_val / capital) ** (1 / years) - 1 if capital > 0 else 0

                summary = pd.DataFrame({
                    "Metric": [
                        "Final Portfolio Value", "Remaining Cash", "Total Invested",
                        "Total Withdrawn", "Net Profit", "CAGR (on capital)"
                    ],
                    "Value": [
                        result_df.iloc[-1]["Portfolio Value"],
                        result_df.iloc[-1]["Remaining Cash"],
                        result_df.iloc[-1]["Total Invested"],
                        result_df.iloc[-1]["Total Withdrawn"],
                        final_val - capital,
                        cagr
                    ]
                })

                with pd.ExcelWriter(RESULT_FILE, engine="openpyxl") as writer:
                    cagr_df.to_excel(writer, sheet_name="Historical Data with CAGR", index=False)
                    result_df.to_excel(writer, sheet_name="Backtesting Results", index=False)
                    summary.to_excel(writer, sheet_name="Performance Summary", index=False)

                return render_template("result.html")

        except Exception as e:
            print("❌ ERROR OCCURRED:", str(e))
            return f"Error: {e}", 500

    return render_template("index.html")


@app.route("/download")
def download():
    return send_file(RESULT_FILE, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
