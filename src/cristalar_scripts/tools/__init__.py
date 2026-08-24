"""Registry of the tool tabs shown on the landing page.

Adding a new tool is one import and one entry in TOOLS - streamlit_app.py
itself does not change.
"""

from cristalar_scripts.tools import faturas_tab

TOOLS = [
    ("Faturas", faturas_tab.render),
]
