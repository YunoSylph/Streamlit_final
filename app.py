import streamlit as st


TIFFANY = "#0ABAB5"

st.set_page_config(
    page_title="Прогнозирование страховых выплат",
    page_icon=None,
    layout="wide",
)

st.markdown(
    f"""
    <style>
    :root {{
        --accent-color: {TIFFANY};
        --accent-dark: #078f8b;
        --ink-color: #172426;
        --muted-color: #5c696b;
        --panel-border: #d7eeee;
        --soft-accent: #e8fbfa;
    }}
    .stApp {{
        color: var(--ink-color);
    }}
    [data-testid="stSidebar"] {{
        border-right: 1px solid var(--panel-border);
    }}
    [data-testid="stMetric"] {{
        background: linear-gradient(180deg, #ffffff 0%, #f7ffff 100%);
        border: 1px solid var(--panel-border);
        border-top: 4px solid var(--accent-color);
        border-radius: 8px;
        padding: 0.85rem 1rem;
        box-shadow: 0 10px 22px rgba(10, 186, 181, 0.08);
    }}
    div.stButton > button[kind="primary"] {{
        background: var(--accent-color);
        border-color: var(--accent-color);
        color: #ffffff;
    }}
    div.stButton > button[kind="primary"]:hover {{
        background: var(--accent-dark);
        border-color: var(--accent-dark);
    }}
    div[data-baseweb="tab-highlight"] {{
        background-color: var(--accent-color);
    }}
    [data-testid="stDataFrame"] {{
        border: 1px solid var(--panel-border);
        border-radius: 8px;
    }}
    h1, h2, h3 {{
        color: var(--ink-color);
    }}
    small, .stCaptionContainer {{
        color: var(--muted-color);
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

pages = {
    "Проект": [
        st.Page("analysis_and_model.py", title="Анализ и модель"),
        st.Page("presentation.py", title="Презентация"),
    ]
}

current_page = st.navigation(pages, position="sidebar", expanded=True)
current_page.run()
