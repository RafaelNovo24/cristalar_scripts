"""Per-user data location, independent of the current working directory.

An installed tool (``uv tool install``) is launched from whatever folder the
user happens to be in, so anything it must find again later - here, the saved
login session - cannot live next to the script.
"""

from pathlib import Path

import platformdirs

DATA_DIR = Path(platformdirs.user_data_dir("cristalar_scripts", appauthor=False))

# Saved browser session (cookies + local storage). While this file exists we
# skip the login step, so we only log in once. Delete it to log in again.
STATE_FILE = DATA_DIR / "auth.json"


def ensure_data_dir() -> None:
    """Create DATA_DIR if it does not exist yet."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
