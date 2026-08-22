"""Fill the PC Plus invoice form from an Excel file.

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
# SELECTORS - the HTML elements of the site, all in one place.
# ---------------------------------------------------------------------------
SELECTORS = {
    # Login page (/Account/Login).
    "login_email_input": "#Email",
    "login_password_input": "#Password",
    "login_entrar_button": "input[type='submit'][value='Entrar']",
    # Home page toolbar.
    "nova_fatura_button": "a[href='/Invoices/Create']",
    # "Tipo de fatura" radio group. The real radio is invisible (iCheck draws
    # its own control on top), so we click the helper overlay instead.
    "tipo_fatura_radio": "#fatura + ins.iCheck-helper",
    # select2 comboboxes: "Cliente" and, in the Items section, "Artigo".
    "cliente_dropdown": "#select2-cliente-container",
    "artigo_dropdown": "#select2-produto-container",
    "dropdown_search_input": "input.select2-search__field",
    # Suggestions of the open dropdown. The "Sem resultados" line is also an
    # <li class="select2-results__option">, but it carries no aria-selected,
    # so requiring that attribute keeps only the real, clickable suggestions.
    "dropdown_options": "li.select2-results__option[aria-selected]",
    # Items section: "Adicionar" creates the item line, which then shows the
    # "Preço Unitário" box. Its id carries a random GUID, so we go by class.
    "adicionar_item_button": "#add-line",
    "unit_price_input": "div.product_total input.product_price",
    # Toolbar button that closes and issues the invoice.
    "finalizar_button": "button[value='Finalizar']",
}

# Item line used in every invoice. The dropdown also holds
# "Prestacao de Servico de Limpeza- ES", so the match must be exact.
ARTIGO = "Prestação de Serviço de Limpeza"

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

# How long to wait for a dropdown to filter its suggestions (milliseconds).
DROPDOWN_WAIT_MS = 2000

# How long to wait for the item line to appear after "Adicionar" (ms).
ITEM_LINE_WAIT_MS = 2000

# Press "Finalizar" at the end. False keeps the run a dry-run: the invoice is
# filled and left open, nothing is issued. Set to True to really submit.
SUBMIT = False


def log_step(message):
    """Print what the script is doing on the page, under the invoice line."""
    print(f"    - {message}")


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


def is_on_login_page(page):
    """True when the site redirected us to the login page."""
    # Going to HOME_URL without a valid session lands on
    # /Account/Login?ReturnUrl=%2fHome%2fIndex, so the URL is enough to tell.
    return LOGIN_PATH.lower() in page.url.lower()


def submit_login_form(page):
    """Fill Email/Password, press Entrar and save the session to STATE_FILE."""
    user, password = get_credentials()
    page.fill(SELECTORS["login_email_input"], user)
    page.fill(SELECTORS["login_password_input"], password)
    page.click(SELECTORS["login_entrar_button"])
    page.wait_for_load_state("networkidle")
    # The login sends us back to ReturnUrl (the home page); still being on the
    # login page means the credentials were refused.
    if is_on_login_page(page):
        raise RuntimeError("Login failed: still on the login page.")
    page.context.storage_state(path=STATE_FILE)
    print("Logged in, session saved.")


def click_nova_fatura_button(page):
    """Go to the home page and press Nova Fatura to open an empty form."""
    page.goto(HOME_URL)
    page.click(SELECTORS["nova_fatura_button"])
    page.wait_for_load_state("networkidle")
    log_step("opened the Nova Fatura form")


def check_tipo_fatura_radio(page):
    """Tick the Fatura option of the Tipo de fatura radio group."""
    page.click(SELECTORS["tipo_fatura_radio"])
    log_step("ticked Tipo de fatura = Fatura")


def search_in_dropdown(page, dropdown, text):
    """Open a select2 dropdown, type text and return the suggestions locator.

    The text is typed key by key: select2 only filters on real keyboard
    events, so setting the value directly would leave the full list on screen.
    """
    page.click(dropdown)
    page.keyboard.type(text, delay=30)
    page.wait_for_timeout(DROPDOWN_WAIT_MS)
    return page.locator(SELECTORS["dropdown_options"])


def select_cliente(page, client):
    """Pick the client in the Cliente dropdown by typing its name.

    Returns True when exactly one suggestion matched and was selected.
    Returns False (with a warning) on zero matches (client does not exist) or
    on several matches (the name is too vague to know which one is meant).
    """
    options = search_in_dropdown(page, SELECTORS["cliente_dropdown"], client)
    count = options.count()

    if count == 0:
        print(f"  WARNING: no client found for '{client}', skipped.")
        return False
    if count > 1:
        names = ", ".join(options.nth(i).inner_text() for i in range(min(count, 5)))
        print(f"  WARNING: {count} clients match '{client}' ({names} ...), "
              "name is not specific enough, skipped.")
        return False

    options.first.click()
    log_step(f"selected Cliente = {client}")
    return True


def select_artigo(page):
    """Pick the fixed item (ARTIGO) in the Artigo dropdown of the Items table."""
    options = search_in_dropdown(page, SELECTORS["artigo_dropdown"], ARTIGO)
    # Typing the full name also matches "... - ES", so take the suggestion
    # whose text is exactly ARTIGO.
    for index in range(options.count()):
        option = options.nth(index)
        if option.inner_text().strip() == ARTIGO:
            option.click()
            log_step(f"selected Artigo = {ARTIGO}")
            return
    raise RuntimeError(f"item '{ARTIGO}' not found in the Artigo dropdown")


def click_adicionar_item_button(page):
    """Press "Adicionar" to create the item line with the price boxes."""
    page.click(SELECTORS["adicionar_item_button"])
    # The line is added by JavaScript, so wait for the price box to show up.
    page.wait_for_selector(SELECTORS["unit_price_input"],
                           timeout=ITEM_LINE_WAIT_MS + 5000)
    page.wait_for_timeout(ITEM_LINE_WAIT_MS)
    log_step("clicked Adicionar, item line created")


def fill_preco_unitario(page, value):
    """Write the amount in the "Preço Unitário" box of the last item line."""
    # The site uses the Portuguese decimal separator (the box starts at
    # "0,0000"), so send the value with a comma.
    amount = f"{value:.2f}".replace(".", ",")
    price_box = page.locator(SELECTORS["unit_price_input"]).last
    price_box.fill(amount)
    # The line total is only recalculated when the box loses focus, so leave
    # it with Tab instead of jumping straight to the next step.
    price_box.press("Tab")
    page.wait_for_timeout(500)
    log_step(f"filled Preço Unitário = {amount}")


def click_finalizar_button(page):
    """Press "Finalizar" to close and issue the invoice."""
    page.click(SELECTORS["finalizar_button"])
    page.wait_for_load_state("networkidle")
    log_step("clicked Finalizar, invoice submitted")


def fill_invoice(context, client, value, index):
    """Fill one invoice in its own tab. True when filled, False when skipped."""
    # One tab per client, all in the same browser window.
    page = context.new_page()
    click_nova_fatura_button(page)
    check_tipo_fatura_radio(page)

    if not select_cliente(page, client):
        page.close()  # nothing to review in this tab
        return False

    select_artigo(page)
    click_adicionar_item_button(page)
    fill_preco_unitario(page, value)

    SCREENSHOTS_DIR.mkdir(exist_ok=True)
    safe_name = "".join(c if c.isalnum() else "_" for c in client)
    shot = SCREENSHOTS_DIR / f"{index:03d}-{safe_name}.png"
    page.screenshot(path=shot)
    log_step(f"screenshot saved to {shot}")

    # The screenshot above is taken before Finalizar, so a dry-run leaves a
    # record of exactly what would have been submitted.
    if SUBMIT:
        click_finalizar_button(page)
    else:
        log_step("dry-run: Finalizar not pressed")
    return True


def main() -> None:
    """Read the Excel file, log in and fill every invoice in dry-run."""
    load_dotenv()
    path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_EXCEL

    invoices = read_invoices(path)
    print(f"{len(invoices)} invoice(s) read from {path}")

    done = 0
    skipped = 0
    failed = 0
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)

        # Reuse the saved session when we have one.
        if STATE_FILE.exists():
            context = browser.new_context(storage_state=STATE_FILE)
        else:
            context = browser.new_context()

        # First tab only checks the session; each invoice gets its own tab.
        login_page = context.new_page()
        login_page.goto(HOME_URL)
        if is_on_login_page(login_page):
            submit_login_form(login_page)
        else:
            print("Already logged in.")
        login_page.close()

        for index, (client, value) in enumerate(invoices, start=1):
            print(f"[{index}] {client} - {value:.2f}")
            # One bad invoice must not stop the remaining ones.
            try:
                if fill_invoice(context, client, value, index):
                    done += 1
                else:
                    skipped += 1
            except Exception as error:  # noqa: BLE001 - keep going
                print(f"  FAILED: {error}")
                failed += 1

        # Keep the filled tabs on screen so they can be reviewed before the
        # window closes.
        input("Press Enter to close the browser...")
        context.close()
        browser.close()

    mode = "submitted" if SUBMIT else "dry-run"
    print(f"End ({mode}): {done} filled, {skipped} skipped, {failed} failed")
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
