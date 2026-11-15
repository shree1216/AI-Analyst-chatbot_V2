import os
import unittest
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from streamlit.testing.v1 import AppTest

from chat_upload import patch_chat_upload
from excel_support import excel_sheet_names
from predictive_analysis import clean_table, find_target, is_full_analysis_request, predict_target, run_full_analysis

ROOT = Path(__file__).resolve().parents[1]
TEST_ENVIRONMENT = {
    name: value for name, value in os.environ.items()
    if not name.upper().startswith(("AZURE_", "SMTP_", "WORKFORCE_"))
}
USER_PROMPT = "clean the data, generate summary and charts with supporting tables and perform predictive analysis"


def telecom_frame(rows: int = 160) -> pd.DataFrame:
    rng = np.random.default_rng(7)
    tenure = rng.integers(1, 61, rows)
    plan = rng.choice(["Prepaid_Basic", "Prepaid_Plus", "Postpaid_Standard"], rows)
    charge = np.round(rng.uniform(15, 90, rows), 2)
    # Short tenure drives churn, so the model must find Tenure_Months as the top driver.
    churn = np.where(tenure + rng.normal(0, 6, rows) < 18, "Yes", "No")
    return pd.DataFrame({
        "Customer_ID": [f"TEL{index:03d}" for index in range(rows)], "Plan_Type": plan, "Network_Type": rng.choice(["4G", "5G"], rows),
        "Tenure_Months": tenure, "Monthly_Charge_USD": charge, "Churn": churn,
    })


class PredictiveAnalysisTests(unittest.TestCase):
    def test_full_analysis_requests_are_recognised_without_taking_other_routes(self):
        for question in (USER_PROMPT, "predict churn", "Which customers are likely to churn?",
                         "clean the data and create charts", "clean the dataset, make a dashboard with supporting tables"):
            with self.subTest(question=question):
                self.assertTrue(is_full_analysis_request(question))
        for question in ("clean this table", "how many customers churned", "show average salary by department",
                         "create a dashboard", "make charts"):
            with self.subTest(question=question):
                self.assertFalse(is_full_analysis_request(question))

    def test_cleaning_keeps_partial_rows_and_logs_every_change(self):
        frame = pd.DataFrame({
            "Customer_ID": ["C1", "C2", "C2", "C3", "C4", None, "C4", "C5", "C6", "C7"],
            "Dept": ["it", "IT ", "IT ", "It", "N/A", None, "Sales", "sales", "-", "Sales"],
            "Charge": ["$1,200", "82,000", "82,000", "", "45.5", None, "abc", "50", "55", "9000000"],
            "Joined": ["2020-01-05", "2021-02-01", "2021-02-01", "null", "bad date", None, "2022-03-01",
                       "2019-01-01", "2018-02-02", "2017-03-03"],
        })
        original = frame.copy(deep=True)
        clean, log = clean_table(frame)
        pd.testing.assert_frame_equal(frame, original)
        # Only the empty row (sheet row 7) and the exact duplicate (row 4) go; repeated ID C4 is flagged, not removed.
        self.assertEqual(clean["Customer_ID"].tolist(), ["C1", "C2", "C3", "C4", "C4", "C5", "C6", "C7"])
        self.assertEqual(clean["Dept"].fillna("").tolist(), ["IT", "IT", "IT", "", "Sales", "Sales", "", "Sales"])
        self.assertEqual(clean["Charge"].fillna(-1).tolist(), [1200.0, 82000.0, -1, 45.5, -1, 50.0, 55.0, 9000000.0])
        self.assertEqual(clean["Joined"].isna().tolist(), [False, False, True, True, False, False, False, False])
        self.assertTrue(pd.api.types.is_datetime64_any_dtype(clean["Joined"]))
        steps = dict(zip(zip(log["Step"], log["Column"]), log["Rows affected"]))
        self.assertEqual(steps[("Removed empty rows", "(all)")], 1)
        self.assertEqual(steps[("Removed duplicate rows", "(all)")], 1)
        self.assertEqual(steps[("Flagged repeated IDs (kept)", "Customer_ID")], 2)
        self.assertEqual(steps[("Set unreadable values to blank", "Charge")], 1)
        self.assertEqual(steps[("Set unreadable values to blank", "Joined")], 1)
        self.assertEqual(steps[("Treated placeholders as blank", "Dept")], 2)
        self.assertEqual(steps[("Flagged extreme outliers (kept)", "Charge")], 1)
        self.assertEqual(steps[("Blanks per column", "Charge")], 2)
        self.assertIn("'it' -> 'IT'", log.set_index(["Step", "Column"]).loc[("Standardised text", "Dept"), "Detail"])
        self.assertIn("row 4 = row 3", log.set_index("Step").loc["Removed duplicate rows", "Detail"])

    def test_prediction_fills_blanks_only_on_its_own_copy(self):
        frame = telecom_frame()
        frame.loc[[3, 9], "Monthly_Charge_USD"] = None
        result = run_full_analysis(USER_PROMPT, frame, "telecom.csv")
        self.assertEqual(int(result.clean["Monthly_Charge_USD"].isna().sum()), 2)
        self.assertEqual(len(result.prediction.predictions), len(frame))
        self.assertIn(f"kept {len(frame)} of {len(frame)} rows", result.answer)
        self.assertTrue(all("Based on" in chart.explanation for chart in result.charts))

    def test_churn_prediction_beats_baseline_and_finds_the_real_driver(self):
        frame = telecom_frame()
        self.assertEqual(find_target(USER_PROMPT, frame), "Churn")
        self.assertEqual(find_target("predict which customers will churn", frame), "Churn")
        result = predict_target(frame, "Churn")
        metrics = dict(zip(result.metrics["Metric"], result.metrics["Value"]))
        self.assertGreater(float(metrics["Accuracy"].rstrip("%")), float(metrics["Baseline accuracy"].rstrip("%")))
        self.assertEqual(result.drivers.iloc[0]["Driver"], "Tenure_Months")
        self.assertIn("less likely 'Yes'", result.drivers.iloc[0]["Effect"])
        self.assertNotIn("Customer_ID", set(result.drivers["Driver"]))
        self.assertEqual(int(result.risk_bands["Records"].sum()), len(frame))
        probability = result.predictions["Probability Churn = Yes"]
        self.assertTrue(probability.is_monotonic_decreasing and probability.between(0, 1).all())

    def test_numeric_target_uses_regression(self):
        frame = telecom_frame()
        frame["Monthly_Charge_USD"] = np.round(20 + frame["Tenure_Months"] * 1.5, 2)
        result = predict_target(frame, find_target("predict Monthly_Charge_USD", frame))
        self.assertEqual((result.task, result.target), ("regression", "Monthly_Charge_USD"))
        self.assertGreater(float(result.metrics.set_index("Metric").loc["R squared", "Value"]), 0.95)

    def test_unknown_target_asks_which_column_to_predict(self):
        frame = telecom_frame().drop(columns="Churn")
        with self.assertRaisesRegex(ValueError, "Which column should I predict"):
            run_full_analysis("perform predictive analysis", frame, "data.csv")

    def test_full_analysis_exports_charts_summary_and_prediction_tabs(self):
        result = run_full_analysis(USER_PROMPT, telecom_frame(), "telecom.csv")
        self.assertEqual(excel_sheet_names(result.excel), [
            "Cleaned Data", "Cleaning Log", "Predictions", "Model Performance", "Key Drivers",
            "Cleaned Data Charts", "Cleaned Data Summary",
        ])
        workbook = load_workbook(BytesIO(result.excel))
        self.addCleanup(workbook.close)
        self.assertEqual(len(workbook["Cleaned Data Charts"]._charts), 4)
        self.assertIn("Key drivers of Churn", [chart.title for chart in result.charts])
        self.assertIn("Predictive analysis", result.answer)


class FullAnalysisUITests(unittest.TestCase):
    def setUp(self):
        self.enterContext(patch("dotenv.load_dotenv"))
        self.enterContext(patch.dict(os.environ, {**TEST_ENVIRONMENT, "WORKFORCE_DATA_MODE": "demo"}, clear=True))
        self.enterContext(patch("socket.create_connection", side_effect=AssertionError("Unexpected network access")))

    def test_combined_request_runs_full_analysis_instead_of_the_calculation_planner(self):
        upload = BytesIO(telecom_frame().to_csv(index=False).encode())
        upload.name = "telecom_dataset.csv"
        self.enterContext(patch_chat_upload(return_value=upload))
        app = AppTest.from_file(str(ROOT / "employee_sql_chatbot.py")).run(timeout=30)
        app.chat_input[0].set_value(USER_PROMPT).run(timeout=60)
        self.assertEqual(len(app.exception), 0)
        message = app.session_state["chat_history"][-1]
        self.assertEqual(message["kind"], "full_analysis")
        self.assertNotIn("error", message)
        response = app.chat_message[-1]
        self.assertEqual([tab.label for tab in response.tabs],
                         ["Charts", "Supporting tables", "Cleaning", "Predictions", "Model"])
        self.assertIn("Download full analysis (Excel)", [button.label for button in response.get("download_button")])
        self.assertIn("Predictions", excel_sheet_names(message["analysis"].excel))


if __name__ == "__main__":
    unittest.main()
