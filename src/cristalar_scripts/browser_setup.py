"""Make sure Playwright's own Chromium build is present before we launch it.

This is a private Chromium download (~150 MB) that Playwright keeps under its
own cache folder. It has nothing to do with the browsers already installed on
the machine and never touches them.
"""

import asyncio
import sys


async def chromium_missing(playwright) -> bool:
    """True when the Chromium build Playwright would launch is not on disk."""
    from pathlib import Path

    return not Path(playwright.chromium.executable_path).exists()


async def install_chromium(on_line=None) -> None:
    """Run ``playwright install chromium``, streaming its output line by line."""
    if getattr(sys, "frozen", False):
        # sys.executable is this exe in a frozen build, not a python
        # interpreter, so "-m playwright" would just relaunch the whole app.
        from cristalar_scripts.launcher import INSTALL_CHROMIUM_FLAG

        argv = [sys.executable, INSTALL_CHROMIUM_FLAG]
    else:
        argv = [sys.executable, "-m", "playwright", "install", "chromium"]

    process = await asyncio.create_subprocess_exec(
        *argv,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.STDOUT,
    )

    lines = []
    async for raw_line in process.stdout:
        line = raw_line.decode(errors="replace").rstrip()
        lines.append(line)
        if on_line is not None:
            on_line(line)

    returncode = await process.wait()
    if returncode != 0:
        tail = "\n".join(lines[-20:])
        raise RuntimeError(f"playwright install chromium failed:\n{tail}")


async def ensure_chromium(playwright, on_line=None) -> None:
    """Install Chromium first when it is missing."""
    if await chromium_missing(playwright):
        if on_line is not None:
            on_line("Chromium not found, downloading it (about 150 MB)...")
        await install_chromium(on_line)
