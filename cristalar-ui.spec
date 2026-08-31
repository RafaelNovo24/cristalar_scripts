# PyInstaller spec: single-file exe for the Streamlit UI.
# Build with:  pyinstaller cristalar-ui.spec

from PyInstaller.utils.hooks import (
    collect_data_files,
    collect_submodules,
    copy_metadata,
)

datas = []
datas += collect_data_files("streamlit")
datas += copy_metadata("streamlit")
datas += collect_data_files("cristalar_scripts")
# streamlit runs this as a script file, so it must exist on disk at runtime,
# not just be importable from the frozen bundle's pyz archive.
datas += [("src/cristalar_scripts/streamlit_app.py", "cristalar_scripts")]

hiddenimports = []
hiddenimports += collect_submodules("streamlit")
hiddenimports += collect_submodules("cristalar_scripts")

a = Analysis(
    ["src/cristalar_scripts/launcher.py"],
    pathex=["src"],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="cristalar-ui",
    console=True,
    onefile=True,
)
