from __future__ import annotations

import altair as alt
import pandas as pd
import streamlit as st

from model_utils import (
    format_currency,
    load_workers_compensation,
    make_prediction_input,
    predict_claim_cost,
    train_models,
)


TIFFANY = "#0ABAB5"
TIFFANY_DARK = "#078F8B"
MODEL_COLORS = {
    "XGBoost": TIFFANY,
    "Random Forest": "#2F80ED",
    "Ridge Regression": "#7A5AF8",
    "Linear Regression": "#F2994A",
}


@st.cache_data(show_spinner=False)
def cached_dataset() -> pd.DataFrame:
    return load_workers_compensation()


@st.cache_resource(show_spinner=False)
def cached_training(sample_size: int | None):
    df = cached_dataset()
    return train_models(df, sample_size=sample_size)


def metric_table(metrics: pd.DataFrame) -> pd.DataFrame:
    table = metrics.copy().reset_index(drop=True)
    table.insert(0, "Место", range(1, len(table) + 1))
    for column in ["MAE", "RMSE"]:
        table[column] = table[column].map(lambda value: f"{value:,.2f}".replace(",", " "))
    table["MSE, млн"] = table["MSE"].map(lambda value: f"{value / 1_000_000:.2f}")
    table = table.drop(columns=["MSE"])
    table["R2"] = table["R2"].map(lambda value: f"{value:.4f}")
    return table


def prediction_sample(predictions: pd.DataFrame, models: list[str], per_model: int = 900) -> pd.DataFrame:
    frames = []
    for model in models:
        model_rows = predictions[predictions["Model"] == model].copy()
        if len(model_rows) > per_model:
            model_rows = model_rows.sample(per_model, random_state=42)
        frames.append(model_rows)

    sampled = pd.concat(frames, ignore_index=True)
    sampled["Residual"] = sampled["Predicted"] - sampled["Actual"]
    sampled["AbsoluteError"] = sampled["Residual"].abs()
    return sampled


def metric_bar_chart(metrics: pd.DataFrame) -> alt.Chart:
    chart_data = metrics.melt(
        id_vars="Model",
        value_vars=["MAE", "RMSE"],
        var_name="Метрика",
        value_name="Значение",
    )
    return (
        alt.Chart(chart_data)
        .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
        .encode(
            x=alt.X("Model:N", title="", sort=list(metrics["Model"])),
            y=alt.Y("Значение:Q", title="Ошибка, $"),
            color=alt.Color(
                "Метрика:N",
                scale=alt.Scale(range=[TIFFANY, "#748184"]),
                legend=alt.Legend(title="Метрика"),
            ),
            tooltip=[
                alt.Tooltip("Model:N", title="Модель"),
                alt.Tooltip("Метрика:N"),
                alt.Tooltip("Значение:Q", format=",.2f"),
            ],
        )
        .properties(height=320)
    )


def importance_chart(importance: pd.DataFrame) -> alt.Chart:
    chart_data = importance.sort_values("Importance")
    return (
        alt.Chart(chart_data)
        .mark_bar(cornerRadiusEnd=4, color=TIFFANY)
        .encode(
            x=alt.X("Importance:Q", title="Важность"),
            y=alt.Y("Feature:N", title="", sort=None),
            tooltip=[
                alt.Tooltip("Feature:N", title="Признак"),
                alt.Tooltip("Importance:Q", title="Важность", format=".4f"),
            ],
        )
        .properties(height=360)
    )


def comparison_chart(data: pd.DataFrame, mode: str, target_cap: float | None) -> alt.Chart:
    model_order = list(MODEL_COLORS.keys())
    color = alt.Color(
        "Model:N",
        scale=alt.Scale(domain=model_order, range=[MODEL_COLORS[name] for name in model_order]),
        legend=None,
    )

    if mode == "Факт vs прогноз":
        limit = float(target_cap or data[["Actual", "Predicted"]].max().max())
        return (
            alt.Chart(data)
            .mark_circle(size=28, opacity=0.42)
            .encode(
                x=alt.X("Actual:Q", title="Реальная стоимость, $", scale=alt.Scale(domain=[0, limit])),
                y=alt.Y("Predicted:Q", title="Прогноз, $", scale=alt.Scale(domain=[0, limit])),
                color=color,
                tooltip=[
                    alt.Tooltip("Model:N", title="Модель"),
                    alt.Tooltip("Actual:Q", title="Факт", format=",.2f"),
                    alt.Tooltip("Predicted:Q", title="Прогноз", format=",.2f"),
                ],
            )
            .facet(column=alt.Column("Model:N", title=""))
            .properties(
                columns=2,
                bounds="flush",
            )
        )

    if mode == "Остатки":
        return (
            alt.Chart(data)
            .mark_circle(size=25, opacity=0.42)
            .encode(
                x=alt.X("Actual:Q", title="Реальная стоимость, $"),
                y=alt.Y("Residual:Q", title="Остаток, $"),
                color=color,
                tooltip=[
                    alt.Tooltip("Model:N", title="Модель"),
                    alt.Tooltip("Actual:Q", title="Факт", format=",.2f"),
                    alt.Tooltip("Residual:Q", title="Остаток", format=",.2f"),
                ],
            )
            .facet(column=alt.Column("Model:N", title=""))
            .properties(columns=2, bounds="flush")
        )

    return (
        alt.Chart(data)
        .mark_boxplot(size=42)
        .encode(
            x=alt.X("Model:N", title="", sort=model_order),
            y=alt.Y("AbsoluteError:Q", title="Абсолютная ошибка, $"),
            color=color,
            tooltip=[
                alt.Tooltip("Model:N", title="Модель"),
                alt.Tooltip("AbsoluteError:Q", title="Абсолютная ошибка", format=",.2f"),
            ],
        )
        .properties(height=360)
    )


def analysis_and_model_page() -> None:
    st.title("Прогнозирование стоимости страховых выплат")
    st.caption("Датасет Workers Compensation, OpenML ID 42876")

    with st.sidebar:
        st.header("Параметры обучения")
        use_full_data = st.toggle("Весь датасет", value=True)
        sample_size = None
        if not use_full_data:
            sample_size = st.slider(
                "Размер выборки",
                min_value=10000,
                max_value=80000,
                step=10000,
                value=30000,
            )
        train_button = st.button("Загрузить данные и обучить модели", type="primary")

    if train_button or "training_result" not in st.session_state:
        with st.spinner("Загрузка данных и обучение моделей..."):
            st.session_state["training_result"] = cached_training(sample_size)
            st.session_state["raw_df"] = cached_dataset()

    result = st.session_state["training_result"]
    df = st.session_state["raw_df"]

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Записей", f"{result.data_shape[0]:,}".replace(",", " "))
    col2.metric("Признаков", result.data_shape[1] - 1)
    col3.metric("Отсечение выбросов", format_currency(float(result.target_cap)))
    col4.metric("Лучшая модель", result.best_model_name)
    col5.metric("RMSE", format_currency(float(result.metrics.iloc[0]["RMSE"])))

    overview_tab, modeling_tab, comparison_tab, prediction_tab = st.tabs(
        ["Данные", "Модели", "Сравнение графиков", "Предсказание"]
    )

    with overview_tab:
        left, right = st.columns([2, 1])
        with left:
            st.subheader("Первые строки датасета")
            st.dataframe(df.head(10), width="stretch")
        with right:
            st.subheader("Целевая переменная")
            summary = pd.DataFrame(
                {
                    "Показатель": ["Минимум", "Медиана", "Среднее", "Максимум"],
                    "Значение": [
                        format_currency(result.target_summary["min"]),
                        format_currency(result.target_summary["median"]),
                        format_currency(result.target_summary["mean"]),
                        format_currency(result.target_summary["max"]),
                    ],
                }
            )
            st.dataframe(summary, hide_index=True, width="stretch")

        st.subheader("Статистика числовых признаков")
        st.dataframe(df.describe().T, width="stretch")

    with modeling_tab:
        left, right = st.columns([1, 1])
        with left:
            st.subheader("Метрики качества")
            st.dataframe(metric_table(result.metrics), hide_index=True, width="stretch")
            st.altair_chart(metric_bar_chart(result.metrics), width="stretch")
        with right:
            st.subheader("Важность признаков")
            st.altair_chart(importance_chart(result.feature_importance), width="stretch")

        best_predictions = result.predictions[result.predictions["Model"] == result.best_model_name]
        st.subheader("Предсказанные и реальные значения")
        best_sample = best_predictions.sample(min(1800, len(best_predictions)), random_state=42)
        st.altair_chart(
            comparison_chart(best_sample, "Факт vs прогноз", result.target_cap),
            width="stretch",
        )

    with comparison_tab:
        selected_models = st.multiselect(
            "Модели",
            options=list(result.metrics["Model"]),
            default=list(result.metrics["Model"]),
        )
        mode = st.segmented_control(
            "Тип графика",
            options=["Факт vs прогноз", "Остатки", "Абсолютные ошибки"],
            default="Факт vs прогноз",
        )
        if not selected_models:
            st.warning("Выберите хотя бы одну модель.")
        else:
            comparison = prediction_sample(result.predictions, selected_models)
            selected_metrics = result.metrics[result.metrics["Model"].isin(selected_models)]
            st.altair_chart(
                comparison_chart(comparison, str(mode), result.target_cap),
                width="stretch",
            )
            st.dataframe(metric_table(selected_metrics), hide_index=True, width="stretch")

    with prediction_tab:
        st.subheader("Новый страховой случай")
        best_model = result.models[result.best_model_name]
        top_descriptions = df["ClaimDescription"].value_counts().head(20).index.tolist()

        with st.form("prediction_form"):
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                age = st.number_input("Возраст", min_value=13, max_value=76, value=35)
                gender = st.selectbox("Пол", ["M", "F"])
                marital_status = st.selectbox("Семейное положение", ["S", "M", "U"])
                dependent_children = st.number_input(
                    "Дети на иждивении", min_value=0, max_value=6, value=0
                )
                dependents_other = st.number_input(
                    "Другие иждивенцы", min_value=0, max_value=3, value=0
                )
            with col_b:
                weekly_pay = st.number_input(
                    "Еженедельная зарплата",
                    min_value=0.0,
                    max_value=5000.0,
                    value=500.0,
                    step=50.0,
                )
                part_time_full_time = st.selectbox("Тип занятости", ["F", "P"])
                hours_worked_per_week = st.number_input(
                    "Часов в неделю", min_value=0, max_value=80, value=38
                )
                days_worked_per_week = st.number_input(
                    "Дней в неделю", min_value=1, max_value=7, value=5
                )
                initial_case_estimate = st.number_input(
                    "Начальная оценка",
                    min_value=1.0,
                    max_value=600000.0,
                    value=5000.0,
                    step=500.0,
                )
            with col_c:
                accident_date = st.date_input("Дата несчастного случая")
                accident_time = st.time_input("Время несчастного случая")
                reported_date = st.date_input("Дата сообщения")
                claim_description = st.selectbox("Описание заявки", top_descriptions)

            submitted = st.form_submit_button("Рассчитать прогноз")

        if submitted:
            accident_datetime = pd.Timestamp.combine(accident_date, accident_time)
            date_reported = pd.Timestamp(reported_date)
            raw_input = make_prediction_input(
                age=age,
                gender=gender,
                marital_status=marital_status,
                dependent_children=dependent_children,
                dependents_other=dependents_other,
                weekly_pay=weekly_pay,
                part_time_full_time=part_time_full_time,
                hours_worked_per_week=hours_worked_per_week,
                days_worked_per_week=days_worked_per_week,
                claim_description=claim_description,
                initial_case_estimate=initial_case_estimate,
                accident_datetime=accident_datetime,
                date_reported=date_reported,
            )
            prediction = predict_claim_cost(best_model, raw_input)
            st.success(f"Прогноз итоговой стоимости возмещения: {format_currency(prediction)}")


analysis_and_model_page()
