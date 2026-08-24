"""Landing page: one tab per tool. Run with `cristalar-ui`."""

import sys

import streamlit as st

from cristalar_scripts.tools import TOOLS

DEFAULT_PORT = 8501


def main() -> None:
    """Render the landing page with one tab per registered tool."""
    st.set_page_config(page_title="Cristalar scripts")
    st.title("Cristalar scripts")

    tabs = st.tabs([label for label, _ in TOOLS])
    for tab, (_, render) in zip(tabs, TOOLS):
        with tab:
            render()


def cli() -> None:
    """Console-script entry point: `cristalar-ui [--port N]`."""
    from streamlit.web import cli as stcli

    port = DEFAULT_PORT
    args = sys.argv[1:]
    if "--port" in args:
        port = args[args.index("--port") + 1]

    sys.argv = [
        "streamlit", "run", __file__,
        "--server.port", str(port),
        "--server.address", "localhost",
    ]
    sys.exit(stcli.main())


if __name__ == "__main__":
    main()
