"""Streamlit tab: download the template, upload it filled in, run the fill."""

import os
import queue
import sys
from io import BytesIO

import streamlit as st

from cristalar_scripts import validar_faturas as vf
from cristalar_scripts.async_runner import AsyncRunner
from cristalar_scripts.criar_template import template_bytes


def _can_show_browser() -> bool:
    """False on a headless Linux server (e.g. Streamlit Cloud) with no X server."""
    return sys.platform != "linux" or bool(os.environ.get("DISPLAY"))


def _state():
    """Session-state keys used by this tab, created once per browser session."""
    if "faturas_runner" not in st.session_state:
        st.session_state.faturas_runner = None
    if "faturas_session" not in st.session_state:
        st.session_state.faturas_session = None
    return st.session_state


def _run_invoices(state, invoices, credentials, headless, submit, parallel_tabs):
    """Start the browser (first run only) and fill every invoice, live."""
    events = queue.Queue()

    if state.faturas_runner is None:
        state.faturas_runner = AsyncRunner()
    runner = state.faturas_runner
    session = state.faturas_session

    async def work():
        nonlocal session
        if session is None:
            session = vf.InvoiceSession()
            await session.start(
                credentials, headless=headless,
                on_event=lambda e: events.put(("log", e.get("log", ""))))
        return await session.fill(
            invoices, submit=submit, parallel_tabs=parallel_tabs,
            on_event=lambda e: events.put(("invoice", e)))

    future = runner.submit(work())

    status = st.status("Running...", expanded=True)
    with status:
        while not future.done():
            try:
                kind, payload = events.get(timeout=0.2)
            except queue.Empty:
                continue
            if kind == "log" and payload:
                st.write(payload)
            elif kind == "invoice":
                with st.expander(f"[{payload['index']}] {payload['client']} "
                                  f"- {payload['value']:.2f}"):
                    for step in payload["steps"]:
                        st.write(step)

        # Drain whatever arrived between the last check and completion.
        while True:
            try:
                kind, payload = events.get_nowait()
            except queue.Empty:
                break
            if kind == "log" and payload:
                st.write(payload)
            elif kind == "invoice":
                with st.expander(f"[{payload['index']}] {payload['client']} "
                                  f"- {payload['value']:.2f}"):
                    for step in payload["steps"]:
                        st.write(step)

        try:
            results = future.result()
        except Exception as error:  # noqa: BLE001 - show it, keep the tab alive
            status.update(label="Failed", state="error")
            st.error(str(error))
            return None
        finally:
            state.faturas_session = session

        status.update(label="Done", state="complete")
    return results


def _close_browser(state):
    if state.faturas_session is not None:
        future = state.faturas_runner.submit(state.faturas_session.close())
        future.result()
        state.faturas_session = None


def render():
    state = _state()

    st.download_button(
        "Download template", template_bytes(), "faturas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    uploaded = st.file_uploader("Upload the filled template", type=["xlsx"])
    invoices = None
    if uploaded is not None:
        try:
            invoices = vf.read_invoices(BytesIO(uploaded.getvalue()))
        except Exception as error:  # noqa: BLE001 - bad file, not a bug
            st.error(f"Could not read the file: {error}")
        else:
            st.caption(f"{len(invoices)} invoice(s) read")
            rows = [{"cliente": client, "valor": value} for client, value in invoices]
            st.dataframe(rows, hide_index=True)

    st.subheader("PC Plus login")
    user = st.text_input("User")
    password = st.text_input("Password", type="password")

    st.subheader("Options")
    submit = st.checkbox("Really submit (Finalizar)", value=False)
    if submit:
        st.warning("This issues the invoices for real.")
    can_show_browser = _can_show_browser()
    show_browser = st.checkbox(
        "Show browser", value=can_show_browser, disabled=not can_show_browser)
    if not can_show_browser:
        st.caption("Not available on this server (no display) — runs headless.")
    parallel_tabs = st.number_input(
        "Parallel tabs", min_value=1, value=vf.PARALLEL_TABS)

    can_run = invoices is not None and user and password
    if st.button("Run", disabled=not can_run):
        results = _run_invoices(
            state, invoices, (user, password),
            headless=not show_browser, submit=submit,
            parallel_tabs=int(parallel_tabs))
        if results is not None:
            done = results.count(True)
            skipped = results.count(False)
            failed = results.count(None)
            mode = "submitted" if submit else "dry-run"
            summary = f"{mode}: {done} filled, {skipped} skipped, {failed} failed"
            if failed:
                st.warning(summary)
            else:
                st.success(summary)

    if state.faturas_session is not None:
        st.button("Close browser", on_click=_close_browser, args=(state,))
