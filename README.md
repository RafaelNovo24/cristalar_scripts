# cristalar_scripts

## What this is

These are scripts for Cristalar's invoicing on PC Plus. The main tool reads a
spreadsheet of clients and amounts and fills the PC Plus invoice form for each
one, one client per browser tab. There is a small web page to run it from, and
a command line for the same job.

## Install

Install `uv` first, if it is not already on the machine. This is a one-time
step per computer:

```powershell
powershell -ExecutionPolicy Bypass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Then install the tool straight from the repository:

```powershell
uv tool install git+https://github.com/RafaelNovo24/cristalar_scripts.git
```

This puts `cristalar-ui`, `cristalar-faturas` and `cristalar-template` on the
PATH, each in their own isolated environment, so they will not clash with
anything else installed on the machine.

The first time a run starts, it downloads its own copy of Chromium (about
150 MB). This is separate from whatever browser is normally used on the
computer - it is not installed as a program, has no icon, and is only used by
this tool to fill in the invoice form.

## Using the web page

Start the page with:

```powershell
cristalar-ui
```

A browser tab opens at `http://localhost:8501` with a "Faturas" tab. The
steps, in order:

1. Press "Download template" to get an empty `faturas.xlsx`.
2. Open it and fill in one row per invoice: `cliente` with the client's name
   and `valor` with the amount.
3. Upload the filled file back into the page.
4. Type the PC Plus user and password. These are only kept for the current
   page session and are never written to disk.
5. Leave "Really submit" unticked for a trial run, and press "Run".

A trial run fills every invoice and leaves the tabs open so they can be
checked by eye. Nothing is sent to PC Plus until "Really submit" is ticked.

## What a run does and does not do

- With "Really submit" off, every invoice is filled and left open in its own
  tab. No invoice is issued.
- With "Really submit" on, each filled invoice is also closed with
  "Finalizar", which issues it for real.
- A client name that matches nothing in PC Plus, or matches more than one
  client, is skipped rather than guessed. The reason is shown next to that
  row.

## Using the command line

The same tool is available without the web page:

```powershell
cristalar-template
cristalar-faturas faturas.xlsx
```

## Where things are kept

Once logged in, the session is cached so the password is not needed again on
the next run. It is stored at:

```
C:\Users\<you>\AppData\Local\cristalar_scripts\auth.json
```

Delete that file to force a fresh login.

## Building a standalone .exe

For a machine without `uv`, build a single-file executable instead:

```powershell
uv sync --group dev
uv run pyinstaller cristalar-ui.spec --noconfirm
```

This produces `dist\cristalar-ui.exe` (~115 MB, self-contained). Double-click
it, or run it from a terminal with `cristalar-ui.exe --port N`, and it opens
the same web page in the default browser automatically.

## Updating and removing

```powershell
uv tool upgrade cristalar-scripts
uv tool uninstall cristalar-scripts
```