"""Full analysis for an uploaded table: clean the data, chart it, and predict a target column.

The model is intentionally small and explainable (NumPy only, no extra dependencies):
- yes/no or category targets (e.g. Churn): multinomial logistic regression
- numeric targets (e.g. Monthly_Charge_USD): ridge linear regression
Accuracy is measured on a 25% hold-out sample, key drivers by permutation importance.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

from data_explorer import (
    ChartResult, ChartSpec, _chart_figure, _find_column, auto_charts, build_charts, categories, column_kind,
    missing_mask, numbers,
)
from excel_support import generate_excel_download
from workbook_charts import apply_excel_dashboard

PREDICTIVE_WORDS = re.compile(
    r"\b(?:predict(?:s|ed|ing|ion|ions|ive)?|forecast(?:s|ing)?|machine\s+learning|ml\s+model|propensity|"
    r"classif(?:y|ication)|regression\s+model|likel(?:y|ihood)\s+to)\b", re.I,
)
CLEAN_WORDS = re.compile(r"\b(?:clean(?:s|ed|ing)?|cleanse|tidy|standardi[sz]e|de-?duplicate)\b", re.I)
VISUAL_WORDS = re.compile(r"\b(?:charts?|graphs?|plots?|visuali[sz]\w*|dashboards?|supporting\s+tables?)\b", re.I)
TARGET_HINTS = ("churn", "attrition", "exited", "left", "target", "label", "outcome", "default", "fraud",
                "converted", "conversion", "cancelled", "canceled", "renewed", "resigned", "status")
POSITIVE_LABELS = {"yes", "y", "true", "1", "churned", "left", "exited", "default", "fraud", "converted", "cancelled"}
HOLDOUT_SHARE = 0.25
MAX_CATEGORIES = 15
SEED = 42


def is_full_analysis_request(question: str) -> bool:
    """Prediction requests, or cleaning combined with charts/tables, need the full analysis workflow."""
    return bool(PREDICTIVE_WORDS.search(question) or (CLEAN_WORDS.search(question) and VISUAL_WORDS.search(question)))


def wants_prediction(question: str) -> bool:
    return bool(PREDICTIVE_WORDS.search(question))


# --------------------------------------------------------------------------- cleaning

def _is_identifier(name: str, series: pd.Series) -> bool:
    present = series[~missing_mask(series)]
    named = bool(re.search(r"(?:^|[\s_])id$|^id(?:[\s_]|$)|identifier|(?:^|[\s_])(?:code|number|no)$", name, re.I))
    unique = len(present) > 0 and present.astype(str).nunique() >= 0.95 * len(present)
    return named and unique


def clean_table(frame: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Return the cleaned table and a log of every cleaning step (step, column, rows affected, detail)."""
    clean = frame.copy().reset_index(drop=True)
    clean.columns = [re.sub(r"\s+", " ", str(column)).strip() for column in clean.columns]
    log: list[dict[str, object]] = []

    def note(step: str, column: str, rows: int, detail: str) -> None:
        if rows:
            log.append({"Step": step, "Column": column, "Rows affected": int(rows), "Detail": detail})

    blank = clean.apply(missing_mask).all(axis=1)
    note("Removed empty rows", "(all)", int(blank.sum()), "Rows with no values at all")
    clean = clean[~blank].reset_index(drop=True)

    for column in clean.columns:
        series = clean[column]
        kind = column_kind(series)
        if kind in {"whole number", "decimal number"}:
            converted = numbers(series)
            changed = int((~missing_mask(series) & series.map(lambda value: isinstance(value, str))).sum())
            note("Converted text to numbers", column, changed, "Removed currency symbols, commas and spaces")
            clean[column] = converted
        elif kind in {"text", "true/false"} and not _is_identifier(column, series):
            standard = categories(series)
            present = ~missing_mask(series)
            changed = int((standard[present].astype(str) != series[present].astype(str)).sum())
            note("Standardised text", column, changed, "Trimmed spaces and unified capitalisation/spelling")
            clean[column] = standard
        elif kind == "text":
            trimmed = series.map(lambda value: re.sub(r"\s+", " ", value).strip() if isinstance(value, str) else value)
            changed = series.map(lambda value: isinstance(value, str)) & (trimmed.astype(str) != series.astype(str))
            note("Trimmed text", column, int(changed.sum()), "Removed extra spaces")
            clean[column] = trimmed

    duplicates = clean.astype(str).duplicated()
    note("Removed duplicate rows", "(all)", int(duplicates.sum()), "Exact repeats of an earlier row")
    clean = clean[~duplicates].reset_index(drop=True)
    identifier = next((column for column in clean.columns if _is_identifier(column, clean[column])), None)
    if identifier is not None:
        repeated = clean[identifier].notna() & clean[identifier].astype(str).str.strip().str.upper().duplicated()
        note("Removed duplicate IDs", identifier, int(repeated.sum()), "Kept the first record for each ID")
        clean = clean[~repeated].reset_index(drop=True)

    for column in clean.columns:
        missing = missing_mask(clean[column])
        if not missing.any() or (identifier is not None and column == identifier):
            continue
        if pd.api.types.is_numeric_dtype(clean[column]):
            median = float(clean.loc[~missing, column].median()) if (~missing).any() else 0.0
            clean.loc[missing, column] = median
            note("Filled missing values", column, int(missing.sum()), f"Used the median ({median:,.2f})")
        else:
            clean[column] = clean[column].astype(object)
            clean.loc[missing, column] = "Unknown"
            note("Filled missing values", column, int(missing.sum()), "Marked as 'Unknown'")

    for column in clean.columns:
        values = clean[column]
        if pd.api.types.is_float_dtype(values) and values.notna().all() and (values % 1 == 0).all():
            clean[column] = values.astype("int64")

    if not log:
        log.append({"Step": "No changes needed", "Column": "(all)", "Rows affected": 0,
                    "Detail": "No blanks, duplicates or inconsistent values were found"})
    return clean, pd.DataFrame(log, columns=["Step", "Column", "Rows affected", "Detail"])


# --------------------------------------------------------------------------- target and features

def find_target(question: str, frame: pd.DataFrame) -> str | None:
    columns = [str(column) for column in frame.columns]
    match = re.search(
        r"\b(?:predict(?:s|ing|ion|ions|ive)?(?:\s+analysis)?|forecast|target|classify)\s+(?:of\s+|for\s+|on\s+|the\s+|whether\s+)*"
        r"(?P<phrase>[\w\s-]+?)(?=\s+(?:using|from|based|with|for|by|and|in|on)\b|[,.;?!]|$)",
        question, re.I,
    )
    if match:
        phrase = match["phrase"].replace("_", " ").strip()
        for candidate in (phrase, *phrase.split()):
            found = _find_column(candidate, columns)
            if found is not None and not _is_identifier(found, frame[found]):
                return found
    mentioned = [column for column in columns
                 if re.search(rf"\b{re.escape(column.replace('_', ' '))}\b", question.replace("_", " "), re.I)]
    hinted = [column for column in mentioned if any(hint in column.casefold() for hint in TARGET_HINTS)]
    if len(hinted) == 1:
        return hinted[0]
    for hint in TARGET_HINTS:
        for column in columns:
            if hint in re.split(r"[\s_]+", column.casefold()) and 2 <= frame[column].nunique(dropna=True) <= 10:
                return column
    return None


def _feature_columns(frame: pd.DataFrame, target: str) -> list[str]:
    features = []
    for column in frame.columns:
        series = frame[column]
        if column == target or _is_identifier(column, series):
            continue
        if pd.api.types.is_numeric_dtype(series):
            if series.nunique() > 1:
                features.append(column)
        elif 1 < series.astype(str).nunique() <= max(MAX_CATEGORIES * 3, 1) and series.astype(str).nunique() < 0.5 * len(series):
            features.append(column)
    return features


@dataclass
class _Encoder:
    numeric: dict[str, tuple[float, float]] = field(default_factory=dict)
    levels: dict[str, list[str]] = field(default_factory=dict)

    def fit(self, frame: pd.DataFrame, features: list[str]) -> "_Encoder":
        for column in features:
            if pd.api.types.is_numeric_dtype(frame[column]):
                values = frame[column].astype(float)
                self.numeric[column] = (float(values.mean()), float(values.std(ddof=0)) or 1.0)
            else:
                self.levels[column] = frame[column].astype(str).value_counts().index[:MAX_CATEGORIES].tolist()
        return self

    def transform(self, frame: pd.DataFrame) -> np.ndarray:
        parts = [np.ones((len(frame), 1))]
        for column, (mean, std) in self.numeric.items():
            parts.append(((frame[column].astype(float).to_numpy() - mean) / std).reshape(-1, 1))
        for column, levels in self.levels.items():
            values = frame[column].astype(str).to_numpy()
            parts.append(np.column_stack([values == level for level in levels]).astype(float))
        return np.hstack(parts)


def _softmax(scores: np.ndarray) -> np.ndarray:
    scores = scores - scores.max(axis=1, keepdims=True)
    exponent = np.exp(scores)
    return exponent / exponent.sum(axis=1, keepdims=True)


def _fit_logistic(x: np.ndarray, y: np.ndarray, classes: int, iterations: int = 600, rate: float = 0.5) -> np.ndarray:
    weights = np.zeros((x.shape[1], classes))
    onehot = np.eye(classes)[y]
    for _ in range(iterations):
        gradient = x.T @ (_softmax(x @ weights) - onehot) / len(x)
        gradient[1:] += 1e-3 * weights[1:]
        weights -= rate * gradient
    return weights


def _fit_ridge(x: np.ndarray, y: np.ndarray) -> np.ndarray:
    penalty = 1e-2 * np.eye(x.shape[1])
    penalty[0, 0] = 0
    return np.linalg.solve(x.T @ x + penalty, x.T @ y)


def _split(labels: np.ndarray, classification: bool) -> tuple[np.ndarray, np.ndarray]:
    rng = np.random.default_rng(SEED)
    test: list[int] = []
    groups = [np.flatnonzero(labels == value) for value in np.unique(labels)] if classification else [np.arange(len(labels))]
    for group in groups:
        shuffled = rng.permutation(group)
        test.extend(shuffled[: max(1, round(len(group) * HOLDOUT_SHARE))] if len(group) > 1 else [])
    test_index = np.array(sorted(test), dtype=int)
    train_index = np.setdiff1d(np.arange(len(labels)), test_index)
    return train_index, test_index


def _auc(positive_scores: np.ndarray, truth: np.ndarray) -> float | None:
    positives, negatives = positive_scores[truth == 1], positive_scores[truth == 0]
    if not len(positives) or not len(negatives):
        return None
    ranks = pd.Series(np.concatenate([positives, negatives])).rank().to_numpy()
    return float((ranks[: len(positives)].sum() - len(positives) * (len(positives) + 1) / 2) / (len(positives) * len(negatives)))


@dataclass
class PredictionResult:
    target: str
    task: str                      # "classification" or "regression"
    positive_label: str | None
    metrics: pd.DataFrame
    drivers: pd.DataFrame
    predictions: pd.DataFrame
    risk_bands: pd.DataFrame | None
    insights: list[str]


def predict_target(frame: pd.DataFrame, target: str) -> PredictionResult:
    data = frame[~missing_mask(frame[target]) & (frame[target].astype(str) != "Unknown")].reset_index(drop=True)
    features = _feature_columns(data, target)
    if not features:
        raise ValueError(f"No usable input columns were found to predict {target}.")
    numeric_target = pd.api.types.is_numeric_dtype(data[target]) and data[target].nunique() > 10
    classification = not numeric_target
    if classification:
        labels = data[target].astype(str)
        classes = sorted(labels.unique(), key=lambda label: (label.casefold() in POSITIVE_LABELS, label))
        if len(classes) < 2:
            raise ValueError(f"{target} has only one value, so there is nothing to predict.")
        if len(classes) > 10:
            raise ValueError(f"{target} has {len(classes)} different values; choose a yes/no, category or numeric target.")
        y = labels.map({label: index for index, label in enumerate(classes)}).to_numpy()
    else:
        classes = []
        y = data[target].astype(float).to_numpy()
    if len(data) < 20:
        raise ValueError(f"At least 20 rows with a known {target} are needed for a reliable prediction.")

    train, test = _split(y, classification)
    encoder = _Encoder().fit(data.iloc[train], features)

    def fit(rows: pd.DataFrame, targets: np.ndarray) -> np.ndarray:
        x = encoder.transform(rows)
        return _fit_logistic(x, targets, len(classes)) if classification else _fit_ridge(x, targets)

    def score(weights: np.ndarray, rows: pd.DataFrame, truth: np.ndarray) -> float:
        x = encoder.transform(rows)
        if classification:
            return float((_softmax(x @ weights).argmax(axis=1) == truth).mean())
        residual = truth - x @ weights
        total = ((truth - truth.mean()) ** 2).sum()
        return float(1 - (residual ** 2).sum() / total) if total else 0.0

    weights = fit(data.iloc[train], y[train])
    test_rows, test_truth = data.iloc[test].reset_index(drop=True), y[test]
    x_test = encoder.transform(test_rows)
    metric_rows: list[tuple[str, str, str]] = []
    if classification:
        probabilities = _softmax(x_test @ weights)
        predicted = probabilities.argmax(axis=1)
        accuracy = float((predicted == test_truth).mean())
        baseline = float(np.bincount(y[train], minlength=len(classes)).max() / len(train))
        metric_rows += [
            ("Accuracy", f"{accuracy:.1%}", "Share of hold-out records predicted correctly"),
            ("Baseline accuracy", f"{baseline:.1%}", "Always guessing the most common value"),
        ]
        if len(classes) == 2:
            positive = 1
            hits = int(((predicted == positive) & (test_truth == positive)).sum())
            precision = hits / max(int((predicted == positive).sum()), 1)
            recall = hits / max(int((test_truth == positive).sum()), 1)
            f1 = 2 * precision * recall / (precision + recall) if precision + recall else 0.0
            auc = _auc(probabilities[:, positive], (test_truth == positive).astype(int))
            metric_rows += [
                ("Precision", f"{precision:.1%}", f"Of records predicted '{classes[1]}', share that really were"),
                ("Recall", f"{recall:.1%}", f"Of real '{classes[1]}' records, share the model found"),
                ("F1 score", f"{f1:.2f}", "Balance of precision and recall (1 = perfect)"),
            ]
            if auc is not None:
                metric_rows.append(("ROC AUC", f"{auc:.2f}", "Ranking quality (0.5 = random, 1 = perfect)"))
    else:
        prediction = x_test @ weights
        error = test_truth - prediction
        r2 = score(weights, test_rows, test_truth)
        metric_rows += [
            ("R squared", f"{r2:.2f}", "Share of variation explained (1 = perfect)"),
            ("Mean absolute error", f"{np.abs(error).mean():,.2f}", f"Average miss, in {target} units"),
            ("Root mean squared error", f"{np.sqrt((error ** 2).mean()):,.2f}", "Penalises large misses"),
        ]
    metric_rows += [("Training rows", f"{len(train):,}", "Used to fit the model"),
                    ("Hold-out rows", f"{len(test):,}", "Kept aside to measure accuracy")]

    # Permutation importance: how much the hold-out score drops when one column is shuffled.
    baseline_score = score(weights, test_rows, test_truth)
    rng = np.random.default_rng(SEED)
    importance = {}
    for column in features:
        shuffled = test_rows.copy()
        drops = []
        for _ in range(5):
            shuffled[column] = rng.permutation(test_rows[column].to_numpy())
            drops.append(baseline_score - score(weights, shuffled, test_truth))
        importance[column] = max(float(np.mean(drops)), 0.0)
    total = sum(importance.values()) or 1.0
    driver_rows = []
    for column in sorted(features, key=lambda name: -importance[name]):
        if pd.api.types.is_numeric_dtype(data[column]) and column in encoder.numeric:
            position = 1 + list(encoder.numeric).index(column)
            if classification:
                # Positive (last) class versus the first class: >0 means higher values raise the chance.
                effect = weights[position, -1] - weights[position, 0]
                direction = f"Higher values -> {'more' if effect > 0 else 'less'} likely '{classes[-1]}'"
            else:
                direction = f"Higher values -> {'higher' if weights[position] > 0 else 'lower'} {target}"
        elif classification:
            rates = (data[target].astype(str) == classes[-1]).groupby(data[column].astype(str)).mean()
            direction = f"Highest '{classes[-1]}' rate: {rates.idxmax()} ({rates.max():.0%})"
        else:
            means = data.groupby(data[column].astype(str))[target].mean()
            direction = f"Highest average {target}: {means.idxmax()} ({means.max():,.2f})"
        driver_rows.append({"Driver": column, "Importance %": round(importance[column] / total * 100, 1), "Effect": direction})
    drivers = pd.DataFrame(driver_rows)

    # Final model on all rows, used to score every record.
    final = fit(data, y)
    x_all = encoder.transform(data)
    predictions = data.copy()
    risk_bands = None
    if classification:
        probabilities = _softmax(x_all @ final)
        predictions[f"Predicted {target}"] = [classes[index] for index in probabilities.argmax(axis=1)]
        if len(classes) == 2:
            chance = probabilities[:, 1]
            predictions[f"Probability {target} = {classes[1]}"] = np.round(chance, 3)
            predictions["Risk band"] = np.select([chance >= 0.6, chance >= 0.3], ["High", "Medium"], "Low")
            predictions = predictions.sort_values(f"Probability {target} = {classes[1]}", ascending=False)
            counts = predictions["Risk band"].value_counts()
            risk_bands = pd.DataFrame({"Risk band": ["High", "Medium", "Low"],
                                       "Records": [int(counts.get(band, 0)) for band in ("High", "Medium", "Low")]})
        else:
            predictions[f"Confidence"] = np.round(probabilities.max(axis=1), 3)
    else:
        predictions[f"Predicted {target}"] = np.round(x_all @ final, 2)
        predictions["Difference"] = np.round(predictions[target] - predictions[f"Predicted {target}"], 2)
    predictions = predictions.reset_index(drop=True)

    insights = []
    top = drivers.head(3)
    if classification:
        accuracy_text = next(value for name, value, _ in metric_rows if name == "Accuracy")
        insights.append(f"The model predicts **{target}** correctly for {accuracy_text} of unseen records "
                        f"(always guessing the most common value would score "
                        f"{next(value for name, value, _ in metric_rows if name == 'Baseline accuracy')}).")
        if len(classes) == 2:
            rate = float((data[target].astype(str) == classes[1]).mean())
            insights.append(f"Current '{classes[1]}' rate: {rate:.1%} of {len(data):,} records.")
            if risk_bands is not None:
                high = int(risk_bands.loc[risk_bands["Risk band"] == "High", "Records"].iloc[0])
                insights.append(f"**{high:,}** records are in the **High** risk band (probability 60% or more) "
                                "and are the best candidates for retention action.")
    else:
        r2_text = next(value for name, value, _ in metric_rows if name == "R squared")
        insights.append(f"The model explains R squared = {r2_text} of the variation in **{target}** on unseen records.")
    if not top.empty:
        insights.append("Key drivers: " + "; ".join(
            f"**{row['Driver']}** ({row['Importance %']:.0f}%) - {row['Effect']}" for _, row in top.iterrows()
        ) + ".")
    insights.append("Predictions show association, not cause; validate them before acting on individual records.")
    return PredictionResult(
        target, "classification" if classification else "regression", classes[1] if len(classes) == 2 else None,
        pd.DataFrame(metric_rows, columns=["Metric", "Value", "Meaning"]), drivers, predictions, risk_bands, insights,
    )


# --------------------------------------------------------------------------- full workflow

@dataclass
class FullAnalysis:
    clean: pd.DataFrame
    cleaning_log: pd.DataFrame
    charts: list[ChartResult]
    prediction: PredictionResult | None
    answer: str
    excel: bytes
    charts_sheet: str
    summary_sheet: str
    notes: list[str] = field(default_factory=list)


def _prediction_charts(result: PredictionResult) -> list[ChartResult]:
    charts = []
    drivers = result.drivers.head(8)[["Driver", "Importance %"]]
    if not drivers.empty and drivers["Importance %"].sum() > 0:
        charts.append(ChartResult(
            f"Key drivers of {result.target}", "bar", drivers,
            f"How much each column helps predict {result.target} (share of total importance).",
            figure=_chart_figure(drivers, "bar"),
        ))
    if result.risk_bands is not None and result.risk_bands["Records"].sum() > 0:
        charts.append(ChartResult(
            f"{result.target} risk bands", "column", result.risk_bands,
            "Records grouped by predicted probability: High 60%+, Medium 30-60%, Low under 30%.",
            figure=_chart_figure(result.risk_bands, "column"),
        ))
    return charts


def run_full_analysis(question: str, frame: pd.DataFrame, source_name: str) -> FullAnalysis:
    if frame.empty:
        raise ValueError("The selected worksheet has no data rows to analyse.")
    clean, cleaning_log = clean_table(frame)
    notes: list[str] = []
    data_charts, problems = build_charts(clean, auto_charts(clean))
    notes.extend(problems)

    prediction = None
    if wants_prediction(question):
        target = find_target(question, clean)
        if target is None:
            raise ValueError(
                "Which column should I predict? Say for example 'predict Churn' using one of these columns: "
                + ", ".join(map(str, clean.columns)) + "."
            )
        prediction = predict_target(clean, target)
    prediction_charts = _prediction_charts(prediction) if prediction else []
    charts = data_charts + prediction_charts

    removed = len(frame) - len(clean)
    lines = [f"Cleaned **{source_name}**: {len(frame):,} rows in, {len(clean):,} clean rows out"
             + (f" ({removed:,} removed)" if removed else "") + f", {int(cleaning_log['Rows affected'].sum()):,} "
             "values fixed. See the Cleaning tab for every step."]
    lines.append(f"Built {len(charts)} charts with their supporting tables.")
    if prediction:
        lines.append("**Predictive analysis**")
        lines.extend(f"- {insight}" for insight in prediction.insights)

    tables: list[tuple[str, pd.DataFrame]] = [("Cleaned Data", clean), ("Cleaning Log", cleaning_log)]
    if prediction:
        tables += [("Predictions", prediction.predictions), ("Model Performance", prediction.metrics),
                   ("Key Drivers", prediction.drivers)]
    workbook = generate_excel_download(tables)
    excel_charts = [chart for chart in data_charts[: 4 - len(prediction_charts)] + prediction_charts
                    if 0 < len(chart.table) <= 500][:4]
    missing = int(sum(missing_mask(frame[column]).sum() for column in frame.columns))
    kpis = [("Rows analysed", f"{len(clean):,}"), ("Values fixed", f"{int(cleaning_log['Rows affected'].sum()):,}"),
            ("Missing values found", f"{missing:,}")]
    if prediction:
        headline = prediction.metrics.iloc[0]
        kpis.append((f"{headline['Metric']} ({prediction.target})", headline["Value"]))
    exported = apply_excel_dashboard(
        workbook, "Cleaned Data", [(chart.title, chart.table, chart.chart_type) for chart in excel_charts],
        source_rows=len(clean), source_columns=len(clean.columns), missing_values=missing, kpis=kpis,
    )
    lines.append(f"The Excel download has **Cleaned Data**, **{exported.sheet}**, **{exported.tables_sheet}**, "
                 "**Cleaning Log**" + (", **Predictions**, **Model Performance** and **Key Drivers**" if prediction else "")
                 + " tabs.")
    return FullAnalysis(clean, cleaning_log, charts, prediction, "\n\n".join(lines), exported.content,
                        exported.sheet, exported.tables_sheet or "", notes)
