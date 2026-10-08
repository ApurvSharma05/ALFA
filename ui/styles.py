"""
ui/styles.py

Minimal custom CSS for ALFA Streamlit application.
Theme colors, fonts, and borders are handled by .streamlit/config.toml.
This file only injects minor layout tweaks if needed.
"""


def inject_styles():
    """Inject minimal custom CSS. Theme is handled by config.toml."""
    import streamlit as st
    # No custom CSS needed — all styling is handled by config.toml theming
    # and inline styles in st.html() calls. This function is kept as a
    # hook for any future minor CSS tweaks.
    pass
