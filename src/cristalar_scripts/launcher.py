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

DEFAULT_PORT = 8506

if getattr(sys, "frozen", False):
    # Outside a frozen build, Playwright resolves browsers to the shared
    # %LOCALAPPDATA%\ms-playwright cache. Inside one, its driver has no
    # node_modules ancestor to detect a normal install from, so it silently
    # falls back to a ".local-browsers" folder next to itself instead -
    # different for every run (a fresh _MEI temp dir), so a chromium install
    # from one run is invisible to the next. Pin the path explicitly so
    # install and launch always agree.
    os.environ.setdefault(
        "PLAYWRIGHT_BROWSERS_PATH",
        os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "ms-playwright"),
    )


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


INSTALL_CHROMIUM_FLAG = "--internal-install-chromium"


def main() -> None:
    """Launch the bundled Streamlit app and open it in the default browser."""
    args = sys.argv[1:]

    if INSTALL_CHROMIUM_FLAG in args:
        # Re-exec target used by browser_setup.install_chromium(): in a frozen
        # build, sys.executable is this exe, not a python interpreter, so
        # "-m playwright install" would just relaunch the whole app instead.
        from playwright.__main__ import main as playwright_main

        sys.argv = ["playwright", "install", "chromium"]
        sys.exit(playwright_main())

    from streamlit.web import cli as stcli

    port = DEFAULT_PORT
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
