from pathlib import Path

import streamlit as st

from dashboard_vih_pages import get_navigation_pages


BASE_DIR = Path(__file__).resolve().parent


st.set_page_config(
    page_title="Dashboard REM VIH 2025",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    :root {
        --vih-primary: #005b8c;
        --vih-secondary: #0f80a6;
        --vih-accent: #e45461;
        --vih-ink: #15364c;
        --vih-soft: #eef6fb;
        --vih-border: #c9dcea;
    }
    .stApp {
        background:
            radial-gradient(circle at top left, rgba(15, 128, 166, 0.12), transparent 28%),
            linear-gradient(180deg, #f6fbfe 0%, #ffffff 18%, #ffffff 100%);
    }
    .stApp h1, .stApp h2, .stApp h3 {
        color: var(--vih-primary);
        font-weight: 750;
    }
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    .hero-panel {
        background: linear-gradient(135deg, #0f4c75 0%, #0f80a6 52%, #4aa7c5 100%);
        border-radius: 22px;
        padding: 1.45rem 1.5rem;
        color: #ffffff;
        box-shadow: 0 18px 42px rgba(15, 76, 117, 0.18);
        margin-bottom: 1rem;
    }
    .hero-title {
        font-size: 1.7rem;
        font-weight: 800;
        margin-bottom: 0.45rem;
        letter-spacing: -0.01em;
    }
    .hero-copy {
        font-size: 0.99rem;
        line-height: 1.55;
        opacity: 0.97;
        margin: 0;
    }
    .info-card {
        background: linear-gradient(180deg, #ffffff 0%, #f8fbfe 100%);
        border: 1px solid var(--vih-border);
        border-radius: 16px;
        padding: 1.1rem 1.15rem;
        box-shadow: 0 12px 28px rgba(21, 54, 76, 0.07);
    }
    .info-card-title {
        color: var(--vih-ink);
        font-size: 1.02rem;
        font-weight: 760;
        margin-bottom: 0.45rem;
    }
    .info-card-copy {
        color: #33566f;
        line-height: 1.5;
        margin: 0;
        font-size: 0.96rem;
    }
    .year-badge {
        display: inline-block;
        margin: 0.15rem 0 0.75rem 0;
        padding: 0.35rem 0.78rem;
        border-radius: 999px;
        background: #fff4e8;
        border: 1px solid #f0c788;
        color: #915200;
        font-size: 0.9rem;
        font-weight: 720;
    }
    .soft-note {
        color: #4a667b;
        font-size: 0.95rem;
        line-height: 1.5;
    }
    div[data-testid="stDataFrame"] {
        border-radius: 14px;
        overflow: hidden;
        border: 1px solid #d9e8f1;
        box-shadow: 0 8px 24px rgba(21, 54, 76, 0.05);
    }
    div[data-testid="stDownloadButton"] button,
    div[data-testid="stLinkButton"] a {
        border-radius: 999px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

logo_path = BASE_DIR / "assets" / "seremi_sidebar_logo.svg"
icon_path = BASE_DIR / "assets" / "seremi_sidebar_icon.svg"
if logo_path.exists() and icon_path.exists():
    st.logo(str(logo_path), size="large", icon_image=str(icon_path))

navigation = st.navigation(get_navigation_pages(), position="sidebar", expanded=True)
navigation.run()
