"""
ui/styles.py

Custom CSS for the ALFA Streamlit application.
Injected once via st.markdown at app startup.
"""

CUSTOM_CSS = """
<style>
/* ----------------------------------------------------------------
   ALFA — Custom Streamlit Theme
   ---------------------------------------------------------------- */

/* Import modern font */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

/* Global font */
html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

/* ---- Metric Cards ---- */
.metric-card {
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    text-align: center;
    transition: transform 0.2s ease, box-shadow 0.2s ease;
}
.metric-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.3);
}
.metric-card .metric-value {
    font-size: 2.2rem;
    font-weight: 700;
    margin: 0.3rem 0;
    line-height: 1.1;
}
.metric-card .metric-label {
    font-size: 0.82rem;
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    opacity: 0.7;
}

/* Color variants */
.metric-card.total  .metric-value { color: #60a5fa; }
.metric-card.success .metric-value { color: #34d399; }
.metric-card.failed .metric-value { color: #f87171; }
.metric-card.pending .metric-value { color: #fbbf24; }

/* ---- Status Badges ---- */
.status-badge {
    display: inline-block;
    padding: 0.25rem 0.75rem;
    border-radius: 20px;
    font-size: 0.78rem;
    font-weight: 600;
    letter-spacing: 0.03em;
}
.status-badge.pending {
    background: rgba(251, 191, 36, 0.15);
    color: #fbbf24;
    border: 1px solid rgba(251, 191, 36, 0.3);
}
.status-badge.processing {
    background: rgba(96, 165, 250, 0.15);
    color: #60a5fa;
    border: 1px solid rgba(96, 165, 250, 0.3);
    animation: pulse 1.5s ease-in-out infinite;
}
.status-badge.success {
    background: rgba(52, 211, 153, 0.15);
    color: #34d399;
    border: 1px solid rgba(52, 211, 153, 0.3);
}
.status-badge.failed {
    background: rgba(248, 113, 113, 0.15);
    color: #f87171;
    border: 1px solid rgba(248, 113, 113, 0.3);
}

@keyframes pulse {
    0%, 100% { opacity: 1; }
    50% { opacity: 0.6; }
}

/* ---- Company Status Card ---- */
.company-card {
    background: rgba(255, 255, 255, 0.03);
    border: 1px solid rgba(255, 255, 255, 0.06);
    border-radius: 10px;
    padding: 1rem 1.25rem;
    margin-bottom: 0.6rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.company-card .company-name {
    font-weight: 600;
    font-size: 0.95rem;
}
.company-card .company-mode {
    font-size: 0.75rem;
    opacity: 0.5;
    margin-left: 0.5rem;
}

/* ---- Sidebar polish ---- */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%);
}
section[data-testid="stSidebar"] .stTextInput > label {
    font-weight: 500;
}

/* ---- Primary button ---- */
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
    border: none;
    border-radius: 8px;
    font-weight: 600;
    padding: 0.6rem 1.5rem;
    transition: all 0.2s ease;
}
.stButton > button[kind="primary"]:hover {
    background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
    box-shadow: 0 4px 16px rgba(59, 130, 246, 0.4);
    transform: translateY(-1px);
}

/* ---- Expander ---- */
.streamlit-expanderHeader {
    font-weight: 600;
    font-size: 0.9rem;
}

/* ---- Download button ---- */
.stDownloadButton > button {
    border-radius: 8px;
    font-weight: 500;
}

/* ---- ZIP download highlight ---- */
.zip-download {
    background: linear-gradient(135deg, #059669 0%, #047857 100%);
    border-radius: 10px;
    padding: 1rem;
    margin-top: 1rem;
    text-align: center;
}
</style>
"""


def inject_styles():
    """Call once from app.py to inject custom CSS."""
    import streamlit as st
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
