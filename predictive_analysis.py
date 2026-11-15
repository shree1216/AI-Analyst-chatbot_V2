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
    ChartResult, ChartSpec, _chart_figure, _find_column, auto_charts, build_charts, categories, column_kind, dates,
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

def _is_identifier(name: str, series: pd.Series, share: float = 0.95) -> bool:
    present = series[~missing_mask(series)]
    named = bool(re.search(r"(?:^|[\s_])id$|^id(?:[\s_]|$)|identifier|(?:^|[\s_])(?:code|number|no)$", name, re.I))
    unique = len(present) > 0 and present.astype(str).nunique() >= share * len(present)
    return named and unique


LOG_COLUMNS = ["Step", "Column", "Rows affected", "Detail"]
CHANGE_STEPS = {"Treated placeholders as blank", "Converted text to numbers", "Converted text to dates",
                "Set unreadable values to blank", "Standardised text", "Trimmed text"}
OUTLIER_IQR_MULTIPLIER = 3.0
CONVERT_SHARE = 0.8
MAX_EXAMPLES = 5


def _examples(items: list[str]) -> str:
    shown = ", ".join(items[:MAX_EXAMPLES])
    return shown + (f" (+{len(items) - MAX_EXAMPLES:,} more)" if len(items) > MAX_EXAMPLES else "")


def _row_list(rows: list[int]) -> str:
    return "row" + ("s " if len(rows) > 1 else " ") + _examples([str(row) for row in rows])


def _changes(before: pd.Series, after: pd.Series, mask: pd.Series) -> str:
    pairs = pd.DataFrame({"before": before[mask].map(repr), "after": after[mask].map(_shown)})
    counts = pairs.value_counts(sort=False)
    return _examples([f"{old} -> {new}" + (f" ({count:,}x)" if count > 1 else "") for (old, new), count in counts.items()])


def _shown(value: object) -> str:
    if isinstance(value, pd.Timestamp):
        return value.date().isoformat() if value == value.normalize() else str(value)
    if isinstance(value, float) and value.is_integer():
        return f"{int(value)}"
    return repr(value) if isinstance(value, str) else f"{value}"


def clean_table(frame: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Clean a copy of an uploaded table and log every change (step, column, rows affected, detail).

    Rows are removed only when completely empty or exact duplicates. Blank and unreadable cells stay
    blank so each chart or calculation can skip them only for the columns it uses. Extreme outliers
    are flagged, never removed. The uploaded frame is not modified.
    """
    clean = frame.copy(deep=True).reset_index(drop=True)
    clean.attrs = {}
    clean.columns = [re.sub(r"\s+", " ", str(column)).strip() for column in clean.columns]
    source_rows = list(frame.attrs.get("excel_row_numbers") or [])
    if len(source_rows) != len(frame):
        source_rows = list(range(2, len(frame) + 2))
    rows = pd.Series(source_rows, dtype="int64")
    log: list[dict[str, object]] = []

    def note(step: str, column: str, count: int, detail: str, *, always: bool = False) -> None:
        if count or always:
            log.append({"Step": step, "Column": column, "Rows affected": int(count), "Detail": detail})

    missing = clean.apply(missing_mask) if len(clean.columns) else pd.DataFrame(index=clean.index)
    empty = missing.all(axis=1) if len(clean.columns) else pd.Series(True, index=clean.index)
    note("Removed empty rows", "(all)", int(empty.sum()),
         f"Every cell was blank: {_row_list(rows[empty].tolist())}" if empty.any() else "")
    # Exact duplicates are judged on the values as uploaded; blanks and placeholders count as equal.
    exact = clean.astype(object).where(~missing, "").astype(str)
    duplicate = exact.duplicated() & ~empty
    if duplicate.any():
        first: dict[tuple, int] = {}
        pairs = []
        for index, key in zip(exact.index, exact.itertuples(index=False, name=None)):
            if duplicate[index]:
                pairs.append(f"row {rows[index]} = row {first[key]}")
            elif not empty[index]:
                first.setdefault(key, rows[index])
        note("Removed duplicate rows", "(all)", int(duplicate.sum()), f"Exact repeats of an earlier row: {_examples(pairs)}")
    keep = ~(empty | duplicate)
    clean, rows = clean[keep].reset_index(drop=True), rows[keep].reset_index(drop=True)

    # A lower uniqueness bar here so repeated IDs are still recognised (and flagged) in small tables.
    identifiers = [column for column in clean.columns if _is_identifier(column, clean[column], share=0.8)]
    for column in clean.columns:
        series = clean[column]
        blank = missing_mask(series)
        placeholder = blank & series.map(lambda value: isinstance(value, str))
        if placeholder.any():
            tokens = series[placeholder].map(lambda value: repr(value.strip()) if value.strip() else "''").value_counts()
            note("Treated placeholders as blank", column, int(placeholder.sum()),
                 _examples([f"{token} ({count:,}x)" for token, count in tokens.items()]))
        kind = column_kind(series)
        if kind == "text" and column not in identifiers and (~blank).any():
            # A mostly numeric or date column is converted, and its few unreadable values become blank (rows kept).
            if numbers(series[~blank]).notna().mean() >= CONVERT_SHARE:
                kind = "decimal number"
            elif dates(series[~blank]).notna().mean() >= CONVERT_SHARE:
                kind = "date"
        is_text = series.map(lambda value: isinstance(value, str)) & ~blank
        if kind in {"whole number", "decimal number", "date"}:
            converted = dates(series) if kind == "date" else numbers(series)
            changed = is_text & converted.notna()
            what = "dates" if kind == "date" else "numbers"
            note(f"Converted text to {what}", column, int(changed.sum()), _changes(series, converted, changed))
            unreadable = ~blank & converted.isna()
            note("Set unreadable values to blank", column, int(unreadable.sum()),
                 f"Not a valid {'date' if kind == 'date' else 'number'} at {_row_list(rows[unreadable].tolist())}: "
                 + _examples(series[unreadable].map(repr).tolist()) + "; the rows were kept")
            clean[column] = converted
        elif kind in {"text", "true/false"} and column not in identifiers:
            standard = categories(series)
            changed = ~blank & (standard.astype(str) != series.astype(str))
            note("Standardised text", column, int(changed.sum()),
                 "Trimmed spaces / unified capitalisation: " + _changes(series, standard, changed))
            clean[column] = standard.where(~blank, None)
        else:
            trimmed = series.map(lambda value: re.sub(r"\s+", " ", value).strip() if isinstance(value, str) else value)
            changed = is_text & (trimmed.astype(str) != series.astype(str))
            note("Trimmed text", column, int(changed.sum()), "Removed extra spaces: " + _changes(series, trimmed, changed))
            clean[column] = trimmed.where(~blank, None)
        values = clean[column]
        if pd.api.types.is_float_dtype(values) and len(values) and values.notna().all() and (values % 1 == 0).all():
            clean[column] = values.astype("int64")

    for identifier in identifiers:
        keys = clean[identifier].astype(str).str.strip().str.upper().where(clean[identifier].notna())
        repeated = keys.notna() & keys.duplicated(keep=False)
        if repeated.any():
            note("Flagged repeated IDs (kept)", identifier, int(repeated.sum()),
                 f"Same {identifier} on {_row_list(rows[repeated].tolist())} but other values differ; review them")

    for column in clean.columns:
        values = clean[column]
        if column in identifiers or not pd.api.types.is_numeric_dtype(values) or pd.api.types.is_bool_dtype(values):
            continue
        present = values.dropna().astype(float)
        if len(present) < 5:
            continue
        low_quartile, high_quartile = present.quantile(0.25), present.quantile(0.75)
        spread = high_quartile - low_quartile
        if not spread:
            continue
        low, high = low_quartile - OUTLIER_IQR_MULTIPLIER * spread, high_quartile + OUTLIER_IQR_MULTIPLIER * spread
        outlier = values.notna() & ((values < low) | (values > high))
        if outlier.any():
            found = [f"row {row}: {_shown(float(value))}" for row, value in zip(rows[outlier], values[outlier])]
            note("Flagged extreme outliers (kept)", column, int(outlier.sum()),
                 f"Outside the expected range {_shown(round(float(low), 2))} to {_shown(round(float(high), 2))} "
                 f"(3x interquartile range): {_examples(found)}. Not removed; check them")

    for column in clean.columns:
        blanks = int(clean[column].isna().sum())
        note("Blanks per column", column, blanks,
             f"{blanks:,} of {len(clean):,} rows blank; kept, and skipped only by charts/calculations that use {column}"
             if blanks else "No blanks", always=True)

    if not any(entry["Step"] in CHANGE_STEPS or entry["Step"].startswith("Removed") for entry in log):
        log.insert(0, {"Step": "No changes needed", "Column": "(all)", "Rows affected": 0,
                       "Detail": "No empty or duplicate rows and no inconsistent values were found"})
    clean.attrs["excel_row_numbers"] = rows.tolist()
    return clean, pd.DataFrame(log, columns=LOG_COLUMNS)


def values_changed(cleaning_log: pd.DataFrame) -> int:
    return int(cleaning_log.loc[cleaning_log["Step"].isin(CHANGE_STEPS), "Rows affected"].sum())


def model_input(clean: pd.DataFrame) -> pd.DataFrame:
    """Fill blanks on a separate copy for the prediction model only (median for numbers, 'Unknown' otherwise)."""
    data = clean.copy()
    identifier = next((column for column in data.columns if _is_identifier(column, data[column])), None)
    for column in data.columns:
        blank = missing_mask(data[column])
        if not blank.any() or column == identifier:
            continue
        if pd.api.types.is_numeric_dtype(data[column]):
            data.loc[blank, column] = float(data.loc[~blank, column].median()) if (~blank).any() else 0.0
        else:
            data[column] = data[column].astype(object)
            data.loc[blank, column] = "Unknown"
    for column in data.columns:
        values = data[column]
        if pd.api.types.is_float_dtype(values) and values.notna().all() and (values % 1 == 0).all():
            data[column] = values.astype("int64")
    return data


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


def _prediction_charts(result: PredictionResult, total: int) -> list[ChartResult]:
    charts = []
    used = len(result.predictions)
    basis = f" Based on {used:,} of {total:,} rows (rows with a known {result.target})."
    drivers = result.drivers.head(8)[["Driver", "Importance %"]]
    if not drivers.empty and drivers["Importance %"].sum() > 0:
        charts.append(ChartResult(
            f"Key drivers of {result.target}", "bar", drivers,
            f"How much each column helps predict {result.target} (share of total importance)." + basis,
            figure=_chart_figure(drivers, "bar"),
        ))
    if result.risk_bands is not None and result.risk_bands["Records"].sum() > 0:
        charts.append(ChartResult(
            f"{result.target} risk bands", "column", result.risk_bands,
            "Records grouped by predicted probability: High 60%+, Medium 30-60%, Low under 30%." + basis,
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
        # The model needs complete inputs, so blanks are filled on a separate copy; the cleaned data keeps them.
        prediction = predict_target(model_input(clean), target)
    prediction_charts = _prediction_charts(prediction, len(clean)) if prediction else []
    charts = data_charts + prediction_charts

    steps = cleaning_log.groupby("Step")["Rows affected"].sum()
    removed = {reason: int(steps.get(step, 0)) for reason, step in
               (("completely empty", "Removed empty rows"), ("exact duplicates", "Removed duplicate rows"))}
    reasons = ", ".join(f"{count:,} {reason}" for reason, count in removed.items() if count)
    blanks = int(clean.isna().sum().sum())
    lines = [f"Cleaned a copy of **{source_name}** (the upload is unchanged): kept {len(clean):,} of {len(frame):,} rows"
             + (f"; removed {reasons}" if reasons else "; no rows removed")
             + f". {values_changed(cleaning_log):,} values were standardised or converted; {blanks:,} blank cells were kept "
             "as blank and each chart skips them only for the columns it uses."
             + (f" {int(steps['Flagged extreme outliers (kept)']):,} extreme outliers were flagged, not removed."
                if "Flagged extreme outliers (kept)" in steps else "")
             + " See the Cleaning tab for every change."]
    lines.append(f"Built {len(charts)} charts with their supporting tables; each states the rows it is based on.")
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
    missing = blanks
    kpis = [("Rows analysed", f"{len(clean):,} of {len(frame):,}"), ("Values fixed", f"{values_changed(cleaning_log):,}"),
            ("Blank cells (kept)", f"{missing:,}")]
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
