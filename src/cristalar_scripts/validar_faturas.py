"""Fill the website invoice form from an Excel file.

Dry-run mode: fills the fields and takes a screenshot, never submits.
Usage: python -m cristalar_scripts.validar_faturas [faturas.xlsx]
"""

import os
import sys
from getpass import getpass
from pathlib import Path

from dotenv import load_dotenv
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

# ---------------------------------------------------------------------------
# SELECTORS - replace with the real HTML elements of the site.
# Any Playwright selector works (CSS, "text=...", "#id", etc).
# ---------------------------------------------------------------------------
SELECTORS = {
    "login_user": "#Email",                             # email field
    "login_password": "#Password",                      # password field
    "login_submit": "input[type='submit'][value='Entrar']",  # "Entrar" button
    "new_invoice": "a[href='/Invoices/Create']",     # "Nova Fatura" button
    "field_client": "input[name='client']",      # client field of the form
    "field_value": "input[name='value']",        # value field of the form
    # "submit": ...  <- add when we move past dry-run
}

# Site address - always the same, so it lives here and not in .env.
HOME_URL = "http://www.pcplusonline.pt/Home/Index"

# Without a valid session the site redirects to this path.
LOGIN_PATH = "/Account/Login"

# Excel file used when no argument is given.
DEFAULT_EXCEL = "faturas.xlsx"

# Saved browser session (cookies + local storage). While this file exists we
# skip the login step, so we only log in once. Delete it to log in again.
STATE_FILE = Path("auth.json")

# Folder holding one screenshot per filled invoice.
SCREENSHOTS_DIR = Path("screenshots")

# Visible browser so we can follow the filling.
HEADLESS = False


def read_invoices(path):
    """Read the Excel file and return a list of (client, value) tuples."""
    wb = load_workbook(path, data_only=True)
    ws = wb.active

    invoices = []
    # min_row=2 skips the header row (cliente / valor).
    for client, value in ws.iter_rows(min_row=2, max_col=2, values_only=True):
        if client is None and value is None:
            continue  # empty row
        invoices.append((str(client).strip(), float(value)))
    return invoices


def get_credentials():
    """Return (user, password) from .env, asking the user for what is missing."""
    user = os.getenv("CRISTALAR_USER") or input("User: ")
    password = os.getenv("CRISTALAR_PASSWORD") or getpass("Password: ")
    return user, password


def needs_login(page):
    """True when the site redirected us to the login page."""
    # Going to HOME_URL without a valid session lands on
    # /Account/Login?ReturnUrl=%2fHome%2fIndex, so the URL is enough to tell.
    return LOGIN_PATH.lower() in page.url.lower()


def login(page):
    """Log in with the credentials and save the session to STATE_FILE."""
    user, password = get_credentials()
    page.fill(SELECTORS["login_user"], user)
    page.fill(SELECTORS["login_password"], password)
    page.click(SELECTORS["login_submit"])
    # Wait for the post-login navigation to settle.
    page.wait_for_load_state("networkidle")
    # The login form sends us back to ReturnUrl (the home page); if we are
    # still on the login page the credentials were refused.
    if LOGIN_PATH.lower() in page.url.lower():
        raise RuntimeError("Login failed: still on the login page.")
    # Keep the session so the next runs skip the login.
    page.context.storage_state(path=STATE_FILE)
    print("Logged in, session saved.")


def open_new_invoice(page):
    """Click the "Nova Fatura" button to open the invoice form."""
    page.click(SELECTORS["new_invoice"])
    page.wait_for_load_state("networkidle")


def fill_invoice(page, client, value, index):
    """Fill client and value and save a screenshot. Does not submit."""
    page.fill(SELECTORS["field_client"], client)
    page.fill(SELECTORS["field_value"], f"{value:.2f}")

    SCREENSHOTS_DIR.mkdir(exist_ok=True)
    safe_name = "".join(c if c.isalnum() else "_" for c in client)
    page.screenshot(path=SCREENSHOTS_DIR / f"{index:03d}-{safe_name}.png")

    # DRY-RUN: page.click(SELECTORS["submit"]) plus the wait for the success
    # message would go here once we want to submit for real.


def main() -> None:
    """Read the Excel file, log in and fill every invoice in dry-run."""
    load_dotenv()
    path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_EXCEL

    invoices = read_invoices(path)
    print(f"{len(invoices)} invoice(s) read from {path}")

    done = 0
    failed = 0
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)

        # Reuse the saved session when we have one.
        if STATE_FILE.exists():
            context = browser.new_context(storage_state=STATE_FILE)
        else:
            context = browser.new_context()

        page = context.new_page()
        page.goto(HOME_URL)

        # The site may already have us logged in (saved session or an active
        # one), so only log in when the login form actually shows up.
        if needs_login(page):
            login(page)
        else:
            print("Already logged in.")

        # Open the invoice form.
        open_new_invoice(page)

        for index, (client, value) in enumerate(invoices, start=1):
            # One bad invoice must not stop the remaining ones.
            try:
                fill_invoice(page, client, value, index)
                print(f"[{index}] OK     {client} - {value:.2f}")
                done += 1
            except Exception as error:  # noqa: BLE001 - keep going
                print(f"[{index}] FAILED {client} - {error}")
                failed += 1

        context.close()
        browser.close()

    print(f"End (dry-run): {done} filled, {failed} failed")
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
