"""
ui/styles.py

Minimal custom CSS for ALFA Streamlit application.
Theme colors, fonts, and borders are handled by .streamlit/config.toml.
This file only covers layout patterns that config.toml cannot express.
"""

CUSTOM_CSS = """
<style>
/* ---- Metric Cards (layout only — colors come from theme) ---- */
.metric-card {
    border: 1px solid var(--border-color, rgba(128,128,128,0.2));
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    text-align: center;
    transition: transform 0.2s ease, box-shadow 0.2s ease;
}
.metric-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(0, 0, 0, 0.15);
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

/* Color accents for metric values */
.metric-card.total  .metric-value { color: #60a5fa; }
.metric-card.success .metric-value { color: #34d399; }
.metric-card.failed .metric-value { color: #f87171; }
.metric-card.pending .metric-value { color: #fbbf24; }

/* ---- Company Status Card ---- */
.company-card {
    border: 1px solid var(--border-color, rgba(128,128,128,0.15));
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
</style>
"""


def inject_styles():
    """Inject minimal custom CSS. Theme is handled by config.toml."""
    import streamlit as st
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
