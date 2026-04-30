from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import numpy as np
import pandas as pd
from sklearn.compose import ColumnTransformer
from sklearn.datasets import fetch_openml
from sklearn.ensemble import RandomForestRegressor
from sklearn.linear_model import LinearRegression, Ridge
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from sklearn.model_selection import train_test_split
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OrdinalEncoder, StandardScaler
from xgboost import XGBRegressor


TARGET = "UltimateIncurredClaimCost"
DATE_COLUMNS = ["DateTimeOfAccident", "DateReported"]
LOW_CARDINALITY_CATEGORICAL = ["Gender", "MaritalStatus", "PartTimeFullTime"]
CATEGORICAL_FEATURES = LOW_CARDINALITY_CATEGORICAL + ["ClaimDescription"]
NUMERIC_FEATURES = [
    "Age",
    "DependentChildren",
    "DependentsOther",
    "WeeklyPay",
    "HoursWorkedPerWeek",
    "DaysWorkedPerWeek",
    "InitialCaseEstimate",
    "AccidentYear",
    "AccidentMonth",
    "AccidentDayOfWeek",
    "ReportingDelay",
    "ClaimDescriptionLength",
    "ClaimDescriptionWords",
]
FEATURES = NUMERIC_FEATURES + CATEGORICAL_FEATURES


@dataclass
class TrainingResult:
    metrics: pd.DataFrame
    feature_importance: pd.DataFrame
    predictions: pd.DataFrame
    models: dict[str, Any]
    best_model_name: str
    sample_size: int
    data_shape: tuple[int, int]
    missing_values: dict[str, int]
    target_summary: dict[str, float]
    feature_columns: list[str]
    target_cap: float | None


def load_workers_compensation() -> pd.DataFrame:
    dataset = fetch_openml(data_id=42876, as_frame=True, parser="auto")
    return dataset.frame


def preprocess_data(df: pd.DataFrame) -> pd.DataFrame:
    data = df.copy()
    data["DateTimeOfAccident"] = pd.to_datetime(data["DateTimeOfAccident"], errors="coerce")
    data["DateReported"] = pd.to_datetime(data["DateReported"], errors="coerce")

    data["AccidentYear"] = data["DateTimeOfAccident"].dt.year
    data["AccidentMonth"] = data["DateTimeOfAccident"].dt.month
    data["AccidentDayOfWeek"] = data["DateTimeOfAccident"].dt.dayofweek
    data["ReportingDelay"] = (data["DateReported"] - data["DateTimeOfAccident"]).dt.days
    data["ReportingDelay"] = data["ReportingDelay"].clip(lower=0)

    data["ClaimDescription"] = data["ClaimDescription"].fillna("UNKNOWN").astype(str)
    data["ClaimDescriptionLength"] = data["ClaimDescription"].str.len()
    data["ClaimDescriptionWords"] = data["ClaimDescription"].str.split().str.len()

    for column in LOW_CARDINALITY_CATEGORICAL:
        data[column] = data[column].fillna("UNKNOWN").astype(str)

    for column in NUMERIC_FEATURES + [TARGET]:
        data[column] = pd.to_numeric(data[column], errors="coerce")
        data[column] = data[column].fillna(data[column].median())

    return data[FEATURES + [TARGET]]


def get_preprocessor() -> ColumnTransformer:
    return ColumnTransformer(
        transformers=[
            ("num", StandardScaler(), NUMERIC_FEATURES),
            (
                "cat",
                OrdinalEncoder(handle_unknown="use_encoded_value", unknown_value=-1),
                CATEGORICAL_FEATURES,
            ),
        ],
        remainder="drop",
    )


def get_regressors() -> dict[str, Any]:
    return {
        "Linear Regression": LinearRegression(),
        "Ridge Regression": Ridge(alpha=1.0),
        "Random Forest": RandomForestRegressor(
            n_estimators=80,
            max_depth=18,
            min_samples_leaf=3,
            random_state=42,
            n_jobs=-1,
        ),
        "XGBoost": XGBRegressor(
            n_estimators=180,
            learning_rate=0.08,
            max_depth=5,
            subsample=0.9,
            colsample_bytree=0.9,
            objective="reg:squarederror",
            tree_method="hist",
            random_state=42,
            n_jobs=-1,
        ),
    }


def build_model(regressor: Any) -> Pipeline:
    return Pipeline(
        steps=[
            ("preprocessor", get_preprocessor()),
            ("model", regressor),
        ]
    )


def _rmse(y_true: pd.Series, y_pred: np.ndarray) -> float:
    return float(np.sqrt(mean_squared_error(y_true, y_pred)))


def evaluate_predictions(y_true: pd.Series, y_pred: np.ndarray) -> dict[str, float]:
    clipped_pred = np.maximum(y_pred, 0)
    return {
        "MAE": float(mean_absolute_error(y_true, clipped_pred)),
        "MSE": float(mean_squared_error(y_true, clipped_pred)),
        "RMSE": _rmse(y_true, clipped_pred),
        "R2": float(r2_score(y_true, clipped_pred)),
    }


def train_models(
    df: pd.DataFrame,
    sample_size: int | None = None,
    random_state: int = 42,
    test_size: float = 0.2,
    target_clip_quantile: float | None = 0.95,
) -> TrainingResult:
    prepared = preprocess_data(df)
    if sample_size and sample_size < len(prepared):
        prepared = prepared.sample(sample_size, random_state=random_state)

    target_cap = None
    if target_clip_quantile:
        target_cap = float(prepared[TARGET].quantile(target_clip_quantile))
        prepared[TARGET] = prepared[TARGET].clip(upper=target_cap)

    X = prepared.drop(columns=[TARGET])
    y = prepared[TARGET]
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=test_size, random_state=random_state
    )

    fitted_models: dict[str, Any] = {}
    metric_rows: list[dict[str, float | str]] = []
    prediction_frames: list[pd.DataFrame] = []

    for name, regressor in get_regressors().items():
        model = build_model(regressor)
        model.fit(X_train, y_train)
        y_pred = np.maximum(model.predict(X_test), 0)
        metrics = evaluate_predictions(y_test, y_pred)
        metric_rows.append({"Model": name, **metrics})
        prediction_frames.append(
            pd.DataFrame(
                {
                    "Model": name,
                    "Actual": y_test.to_numpy(),
                    "Predicted": y_pred,
                }
            )
        )
        fitted_models[name] = model

    metrics_df = pd.DataFrame(metric_rows).sort_values("RMSE").reset_index(drop=True)
    best_model_name = str(metrics_df.iloc[0]["Model"])
    best_model = fitted_models[best_model_name]
    feature_importance = get_feature_importance(best_model).head(15)

    return TrainingResult(
        metrics=metrics_df,
        feature_importance=feature_importance,
        predictions=pd.concat(prediction_frames, ignore_index=True),
        models=fitted_models,
        best_model_name=best_model_name,
        sample_size=len(prepared),
        data_shape=tuple(df.shape),
        missing_values={column: int(value) for column, value in df.isna().sum().items()},
        target_summary={
            "min": float(df[TARGET].min()),
            "median": float(df[TARGET].median()),
            "mean": float(df[TARGET].mean()),
            "max": float(df[TARGET].max()),
        },
        feature_columns=list(X.columns),
        target_cap=target_cap,
    )


def get_feature_importance(model: Pipeline) -> pd.DataFrame:
    regressor = model.named_steps["model"]
    preprocessor = model.named_steps["preprocessor"]
    raw_names = preprocessor.get_feature_names_out()
    feature_names = [name.replace("num__", "").replace("cat__", "") for name in raw_names]

    if hasattr(regressor, "feature_importances_"):
        importance = regressor.feature_importances_
    elif hasattr(regressor, "coef_"):
        importance = np.abs(np.ravel(regressor.coef_))
    else:
        importance = np.zeros(len(feature_names))

    return (
        pd.DataFrame({"Feature": feature_names, "Importance": importance})
        .sort_values("Importance", ascending=False)
        .reset_index(drop=True)
    )


def make_prediction_input(
    age: int,
    gender: str,
    marital_status: str,
    dependent_children: int,
    dependents_other: int,
    weekly_pay: float,
    part_time_full_time: str,
    hours_worked_per_week: int,
    days_worked_per_week: int,
    claim_description: str,
    initial_case_estimate: float,
    accident_datetime: pd.Timestamp,
    date_reported: pd.Timestamp,
) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "DateTimeOfAccident": accident_datetime,
                "DateReported": date_reported,
                "Age": age,
                "Gender": gender,
                "MaritalStatus": marital_status,
                "DependentChildren": dependent_children,
                "DependentsOther": dependents_other,
                "WeeklyPay": weekly_pay,
                "PartTimeFullTime": part_time_full_time,
                "HoursWorkedPerWeek": hours_worked_per_week,
                "DaysWorkedPerWeek": days_worked_per_week,
                "ClaimDescription": claim_description,
                "InitialCaseEstimate": initial_case_estimate,
                TARGET: 0.0,
            }
        ]
    )


def predict_claim_cost(model: Pipeline, raw_input: pd.DataFrame) -> float:
    prepared = preprocess_data(raw_input).drop(columns=[TARGET])
    return float(np.maximum(model.predict(prepared), 0)[0])


def format_currency(value: float) -> str:
    return f"${value:,.2f}".replace(",", " ")
