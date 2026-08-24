"""A background asyncio event loop that outlives a single Streamlit rerun.

Streamlit reruns the whole script on every widget interaction, and
`asyncio.run()` tears its loop down when the coroutine finishes. Playwright's
async API needs one loop that stays alive for as long as the browser session
does, so we run one in a dedicated daemon thread and hand it work from
whichever rerun needs it.
"""

import asyncio
import threading


class AsyncRunner:
    """Runs coroutines on a single background event loop."""

    def __init__(self):
        self._loop = asyncio.new_event_loop()
        self._thread = threading.Thread(target=self._loop.run_forever, daemon=True)
        self._thread.start()

    def submit(self, coro):
        """Schedule coro on the background loop, returning a concurrent Future."""
        return asyncio.run_coroutine_threadsafe(coro, self._loop)

    def shutdown(self):
        """Stop the loop and wait for its thread to exit."""
        self._loop.call_soon_threadsafe(self._loop.stop)
        self._thread.join()
