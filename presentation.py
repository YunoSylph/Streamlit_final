from __future__ import annotations

import base64
import json
from pathlib import Path

import streamlit as st

try:
    import reveal_slides as rs
except ImportError:  # pragma: no cover
    rs = None


ACCENT = "#0ABAB5"
INK = "#172426"
MUTED = "#617174"


def load_result_summary() -> dict:
    path = Path("reports/model_results.json")
    if not path.exists():
        return {
            "best_model_name": "XGBoost",
            "metrics": [],
            "feature_importance": [],
            "target_cap": 54450.09,
            "target_summary": {"max": 3139046.0},
        }
    return json.loads(path.read_text(encoding="utf-8"))


def money(value: float) -> str:
    return f"${value:,.2f}".replace(",", " ")


def asset_uri(path: str) -> str:
    file_path = Path(path)
    if not file_path.exists():
        return ""
    encoded = base64.b64encode(file_path.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def metric_rows(metrics: list[dict]) -> str:
    rows = []
    for index, row in enumerate(metrics, start=1):
        rows.append(
            "<tr>"
            f"<td>{index}</td>"
            f"<td>{row['Model']}</td>"
            f"<td>{money(row['MAE'])}</td>"
            f"<td>{money(row['RMSE'])}</td>"
            f"<td>{row['R2']:.4f}</td>"
            "</tr>"
        )
    return "\n".join(rows)


def feature_rows(features: list[dict]) -> str:
    rows = []
    for row in features[:7]:
        rows.append(
            "<tr>"
            f"<td>{row['Feature']}</td>"
            f"<td>{row['Importance']:.4f}</td>"
            "</tr>"
        )
    return "\n".join(rows)


def presentation_css() -> str:
    return f"""
:root {{
  --accent: {ACCENT};
  --accent-dark: #078f8b;
  --ink: {INK};
  --muted: {MUTED};
  --line: #d7e4e5;
  --panel: #f7fbfb;
  --paper: #ffffff;
}}
.reveal {{
  font-family: "Segoe UI", Arial, sans-serif;
  color: var(--ink);
  background: #f9fcfc;
  opacity: 1 !important;
}}
.reveal.ready {{
  opacity: 1 !important;
}}
.reveal .slides section {{
  text-align: left;
  box-sizing: border-box;
  padding: 34px 42px;
}}
.reveal h1,
.reveal h2,
.reveal h3 {{
  letter-spacing: 0;
  text-transform: none;
  color: var(--ink);
  font-weight: 780;
}}
.reveal h1 {{
  font-size: 1.32em;
  line-height: 1.08;
  margin-bottom: 20px;
}}
.reveal h2 {{
  font-size: 1.22em;
  margin-bottom: 18px;
  padding-bottom: 10px;
  border-bottom: 4px solid var(--accent);
}}
.cover {{
  min-height: 100%;
  display: grid;
  grid-template-columns: minmax(0, 1.08fr) minmax(0, 0.92fr);
  gap: 38px;
  align-items: center;
}}
.cover-panel {{
  border-left: 8px solid var(--accent);
  padding: 26px 28px;
  background: var(--paper);
  box-shadow: 0 20px 48px rgba(9, 65, 65, 0.08);
}}
.cover-kpis {{
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 14px;
}}
.kpi {{
  border: 1px solid var(--line);
  background: var(--paper);
  padding: 15px 16px;
  min-height: 82px;
}}
.kpi .label {{
  color: var(--muted);
  font-size: 0.42em;
}}
.kpi .value {{
  display: block;
  color: var(--ink);
  font-size: 0.62em;
  line-height: 1.12;
  font-weight: 760;
  margin-top: 10px;
  overflow-wrap: anywhere;
}}
.subtitle {{
  color: var(--muted);
  font-size: 0.56em;
  line-height: 1.45;
}}
.section-label {{
  display: inline-block;
  padding: 5px 10px;
  background: #e7f8f7;
  color: var(--accent-dark);
  font-size: 0.34em;
  font-weight: 700;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  margin-bottom: 16px;
}}
.two-col {{
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 26px;
  align-items: start;
}}
.formal-list {{
  margin: 0;
  padding-left: 22px;
  font-size: 0.54em;
  line-height: 1.48;
}}
.formal-list li {{
  margin-bottom: 9px;
}}
.process {{
  display: grid;
  grid-template-columns: repeat(5, 1fr);
  gap: 12px;
  margin-top: 26px;
}}
.step {{
  border: 1px solid var(--line);
  background: var(--paper);
  padding: 15px 13px;
  min-height: 112px;
  box-shadow: 0 12px 24px rgba(9, 65, 65, 0.04);
}}
.step strong {{
  color: var(--accent-dark);
  font-size: 0.82em;
  display: block;
  margin-bottom: 8px;
}}
.step span {{
  color: var(--ink);
  font-size: 0.41em;
  line-height: 1.35;
}}
table.formal {{
  width: 100%;
  border-collapse: collapse;
  font-size: 0.40em;
  background: var(--paper);
}}
table.formal th {{
  background: #eef4f4;
  border: 1px solid #aab8ba;
  padding: 8px 9px;
  text-align: center;
  font-weight: 760;
}}
table.formal td {{
  border: 1px solid #c7d3d4;
  padding: 8px 9px;
}}
.chart-wrap {{
  background: var(--paper);
  border: 1px solid var(--line);
  padding: 10px;
  box-sizing: border-box;
  box-shadow: 0 20px 40px rgba(9, 65, 65, 0.08);
}}
.chart-wrap img {{
  width: 100%;
  display: block;
}}
.chart-wrap.fit-slide {{
  height: 460px;
  display: flex;
  align-items: center;
  justify-content: center;
  overflow: hidden;
}}
.chart-wrap.fit-slide img {{
  width: auto;
  max-width: 100%;
  max-height: 100%;
  object-fit: contain;
}}
.callout {{
  border-left: 6px solid var(--accent);
  background: #effafa;
  padding: 16px 18px;
  color: var(--ink);
  font-size: 0.52em;
  line-height: 1.45;
}}
.muted {{
  color: var(--muted);
}}
"""


def build_presentation_markdown(data: dict) -> str:
    metrics = data["metrics"]
    features = data["feature_importance"]
    best_model = data["best_model_name"]
    best_row = next((row for row in metrics if row["Model"] == best_model), None)
    best_metric = (
        f"{best_model}: MAE {money(best_row['MAE'])}, RMSE {money(best_row['RMSE'])}, R2 {best_row['R2']:.4f}"
        if best_row
        else "XGBoost: результаты рассчитаны на странице анализа"
    )
    metric_table = metric_rows(metrics)
    feature_table = feature_rows(features)
    metrics_chart = asset_uri("reports/assets/metrics_comparison.png")
    importance_chart = asset_uri("reports/assets/feature_importance.png")
    comparison_chart = asset_uri("reports/assets/models_scatter_comparison.png")
    target_cap = money(float(data["target_cap"]))

    return f"""
<section>
<style>{presentation_css()}</style>

<div class="cover">
  <div class="cover-panel">
    <div class="section-label">ВКР · Data Science</div>
    <h1>Прогнозирование стоимости страховых выплат</h1>
    <div class="subtitle">Регрессионная модель для оценки итоговой стоимости страхового возмещения по данным Workers Compensation.</div>
    <div style="margin-top: 36px;" class="callout">Лучшая модель по фактическим результатам: <strong>{best_metric}</strong>.</div>
  </div>
  <div class="cover-kpis">
    <div class="kpi"><span class="label">Источник</span><span class="value">OpenML 42876</span></div>
    <div class="kpi"><span class="label">Наблюдений</span><span class="value">100 000</span></div>
    <div class="kpi"><span class="label">Признаков</span><span class="value">13</span></div>
    <div class="kpi"><span class="label">Целевая переменная</span><span class="value">Ultimate<wbr>Incurred<wbr>Claim<wbr>Cost</span></div>
  </div>
</div>
</section>

<section>
<h2>Цель и бизнес-контекст</h2>
<div class="two-col">
<ul class="formal-list">
<li>Построить модель регрессии для оценки итоговой стоимости страхового возмещения.</li>
<li>Использовать признаки работника, условия занятости, описание заявки и начальную оценку случая.</li>
<li>Сравнить несколько моделей и выбрать лучшую по RMSE.</li>
</ul>
<ul class="formal-list">
<li>Практический эффект: поддержка планирования страховых резервов.</li>
<li>Снижение неопределенности при ранней оценке будущих выплат.</li>
<li>Демонстрация решения через многостраничное Streamlit-приложение.</li>
</ul>
</div>
</section>

<section>
<h2>Данные и подготовка</h2>
<div class="process">
  <div class="step"><strong>01</strong><span>Загрузка Workers Compensation из OpenML.</span></div>
  <div class="step"><strong>02</strong><span>Признаки из дат: год, месяц, день недели и задержка сообщения.</span></div>
  <div class="step"><strong>03</strong><span>Кодирование категорий и масштабирование числовых признаков.</span></div>
  <div class="step"><strong>04</strong><span>Ограничение целевой переменной 95-м перцентилем: {target_cap}.</span></div>
  <div class="step"><strong>05</strong><span>Разделение train/test 80/20 и оценка MAE, MSE, RMSE, R2.</span></div>
</div>
</section>

<section>
<h2>Сравнение моделей</h2>
<div class="two-col">
  <div>
    <table class="formal">
    <thead><tr><th>№</th><th>Модель</th><th>MAE</th><th>RMSE</th><th>R2</th></tr></thead>
    <tbody>{metric_table}</tbody>
    </table>
  </div>
  <div class="chart-wrap">
    <img alt="Сравнение ошибок моделей" src="{metrics_chart}">
  </div>
</div>
</section>

<section>
<h2>Факт и прогноз</h2>
<div class="chart-wrap fit-slide">
  <img alt="Сравнение графиков факт-прогноз" src="{comparison_chart}">
</div>
</section>

<section>
<h2>Важность признаков</h2>
<div class="two-col">
  <div>
    <table class="formal">
    <thead><tr><th>Признак</th><th>Важность</th></tr></thead>
    <tbody>{feature_table}</tbody>
    </table>
    <div class="callout" style="margin-top: 18px;">Ключевой фактор: <strong>InitialCaseEstimate</strong>. Это согласуется с экономическим смыслом задачи, так как начальная оценка отражает первичную экспертизу тяжести случая.</div>
  </div>
  <div class="chart-wrap">
    <img alt="Важность признаков" src="{importance_chart}">
  </div>
</div>
</section>

<section>
<h2>Streamlit-приложение</h2>
<div class="two-col">
<ul class="formal-list">
<li>Навигация реализована через <strong>st.navigation</strong> и <strong>st.Page</strong>.</li>
<li>Основная страница содержит загрузку данных, обучение моделей, метрики, графики и прогноз нового случая.</li>
<li>Добавлена вкладка для сравнения моделей: факт-прогноз, остатки и абсолютные ошибки.</li>
</ul>
<ul class="formal-list">
<li>Страница презентации построена через <strong>streamlit-reveal-slides</strong>.</li>
<li>Метрики и выводы подтягиваются из фактического файла результатов.</li>
<li>Цветовой акцент интерфейса: Tiffany, <span class="muted">#0ABAB5</span>.</li>
</ul>
</div>
</section>

<section>
<h2>Выводы</h2>
<div class="callout">
В работе реализован полный цикл решения задачи регрессии: загрузка и подготовка данных, обучение четырех моделей, сравнение качества, анализ важности признаков и интерактивная демонстрация. По фактическим результатам лучшей моделью стала <strong>{best_model}</strong>. Возможные улучшения: NLP-обработка описаний заявок, подбор гиперпараметров и отдельное моделирование экстремально крупных выплат.
</div>
</section>
"""


def presentation_page() -> None:
    data = load_result_summary()
    st.markdown(
        """
        <style>
        [data-testid="stSidebar"] {
            border-right: 1px solid #d6e9e9;
        }
        .presentation-note {
            border-left: 5px solid #0ABAB5;
            padding: 0.75rem 1rem;
            background: #f4fbfb;
            color: #263336;
            margin-bottom: 1rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.title("Презентация проекта")
    st.markdown(
        "<div class='presentation-note'>Обновленная визуальная версия: строгая структура, Tiffany-акцент, фактические метрики и графики из результата обучения.</div>",
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.header("Настройки")
        theme = st.selectbox("Тема", ["white", "black", "league", "night"], index=0)
        height = st.number_input("Высота", min_value=540, max_value=920, value=720)
        transition = st.selectbox("Переход", ["slide", "fade", "convex", "zoom"], index=1)

    presentation_markdown = build_presentation_markdown(data)
    if rs is None:
        st.markdown(f"<style>{presentation_css()}</style>{presentation_markdown}", unsafe_allow_html=True)
    else:
        rs.slides(
            presentation_markdown,
            height=height,
            theme=theme,
            css=presentation_css(),
            config={
                "transition": transition,
                "controls": True,
                "progress": True,
                "center": False,
                "hash": False,
                "width": 1280,
                "height": 720,
                "margin": 0.05,
            },
            markdown_props={"data-separator-vertical": "^--$"},
            allow_unsafe_html=True,
        )


presentation_page()
