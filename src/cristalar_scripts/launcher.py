"""Entry point for the standalone .exe build (PyInstaller).

Runs the Streamlit app the same way `cristalar-ui` does, but from inside a
frozen bundle: Streamlit's script path must be a real file on disk, which
under PyInstaller's onefile mode means the extracted temp dir, not this
launcher's own path.
"""

import os
import sys
import threading
import webbrowser

DEFAULT_PORT = 8501


def _open_browser_when_ready(url: str) -> None:
    import time
    import urllib.request

    for _ in range(100):
        try:
            urllib.request.urlopen(url, timeout=0.5)
            break
        except Exception:
            time.sleep(0.2)
    webbrowser.open(url)


def main() -> None:
    """Launch the bundled Streamlit app and open it in the default browser."""
    from streamlit.web import cli as stcli

    port = DEFAULT_PORT
    args = sys.argv[1:]
    if "--port" in args:
        port = args[args.index("--port") + 1]

    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(__file__))
    app_path = os.path.join(base_dir, "cristalar_scripts", "streamlit_app.py")
    url = f"http://localhost:{port}"

    threading.Thread(target=_open_browser_when_ready, args=(url,), daemon=True).start()

    sys.argv = [
        "streamlit", "run", app_path,
        "--server.port", str(port),
        "--server.address", "localhost",
        "--server.headless", "true",
        "--browser.gatherUsageStats", "false",
        "--global.developmentMode", "false",
    ]
    sys.exit(stcli.main())


if __name__ == "__main__":
    main()
