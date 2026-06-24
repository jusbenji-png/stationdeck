# =============================================================
# stationdeck.spec
# PyInstaller build recipe for StationDeck
#
# Run with:
#   pyinstaller stationdeck.spec
# Or use build.bat which does this automatically.
#
# Output: dist\StationDeck\StationDeck.exe
# =============================================================

from PyInstaller.utils.hooks import collect_data_files, collect_submodules
import sys
from pathlib import Path

# Project root is where this spec file lives
PROJECT_ROOT = Path(SPECPATH)

# -- Data files to bundle -----------------------------------------------------
# These are non-Python files the app needs at runtime.
# Format: (source_path, destination_folder_inside_bundle)
#
# We bundle:
#   - web/templates/          -> web/templates/
#   - web/static/             -> web/static/
#   - config/stations/*.yaml  -> config/stations/
#   - config/version.txt      -> config/
#
# We DO NOT bundle:
#   - data/            <- stays outside, managed by the station
#   - reports/         <- stays outside
#   - logs/            <- stays outside
#   - .env             <- stays outside (contains secrets)
#   - config/license.key <- stays outside (per-machine)

datas = [
    # Flask templates -- must be accessible at runtime
    (str(PROJECT_ROOT / "web" / "templates"), "web/templates"),
    # Static files -- CSS, JS, images
    (str(PROJECT_ROOT / "web" / "static"), "web/static"),
    # Station YAML configs -- bundled as default configs
    (str(PROJECT_ROOT / "config" / "stations"), "config/stations"),
    # Version file -- read by src/updater.py on startup
    (str(PROJECT_ROOT / "config" / "version.txt"), "config"),
]

# -- Hidden imports ------------------------------------------------------------
# PyInstaller's static analysis misses some dynamic imports.
# We list them explicitly here.

hidden_imports = [
    # Flask and its internals
    "flask",
    "flask.templating",
    "werkzeug",
    "werkzeug.serving",
    "werkzeug.debug",
    "jinja2",
    "jinja2.ext",
    "markupsafe",
    "itsdangerous",
    "click",

    # Excel reading
    "openpyxl",
    "openpyxl.styles",
    "openpyxl.utils",
    "openpyxl.reader.excel",
    "openpyxl.writer.excel",
    "openpyxl.workbook",
    "openpyxl.worksheet",

    # PDF generation
    "reportlab",
    "reportlab.platypus",
    "reportlab.lib",
    "reportlab.lib.pagesizes",
    "reportlab.lib.styles",
    "reportlab.lib.colors",
    "reportlab.lib.units",
    "reportlab.lib.enums",
    "reportlab.pdfgen",
    "reportlab.pdfbase",
    "reportlab.pdfbase.ttfonts",
    "reportlab.pdfbase.pdfmetrics",

    # Word documents
    "docx",
    "docx.shared",
    "docx.enum.text",
    "docx.oxml",
    "docx.oxml.ns",

    # Data processing
    "pandas",
    "numpy",

    # Database
    "sqlite3",

    # Encryption / license
    "cryptography",
    "cryptography.fernet",
    "cryptography.hazmat",
    "cryptography.hazmat.primitives",
    "cryptography.hazmat.backends",

    # Config
    "yaml",
    "dotenv",

    # OpenAI
    "openai",
    "httpx",

    # HTTP requests -- used by auth_client and updater
    "requests",
    "urllib3",
    "certifi",
    "charset_normalizer",
    "idna",

    # Imaging -- used by reportlab, openpyxl, and OCR
    "PIL",
    "PIL.Image",
    "PIL.ImageFilter",
    "PIL.ImageOps",

    # OCR -- pytesseract is optional; app degrades gracefully if absent
    "pytesseract",

    # Standard library items PyInstaller sometimes misses
    "email",
    "email.mime",
    "email.mime.multipart",
    "email.mime.text",
    "email.mime.base",
    "smtplib",
    "ssl",
    "json",
    "logging",
    "logging.handlers",
    "threading",
    "webbrowser",
    "winreg",
    "socket",

    # StationDeck src modules -- imported dynamically inside route functions
    # so PyInstaller static analysis may miss them
    "src.reader",
    "src.reader_stock",
    "src.reader_shop",
    "src.reader_manager",
    "src.processor",
    "src.ai_engine",
    "src.exporter",
    "src.emailer",
    "src.database",
    "src.license",
    "src.ocr_reader",
    "src.daily_entry_processor",
    "src.auth_client",       # Phase 14B — auth server communication
    "src.updater",           # Auto-update checker
    "src.capture_template_builder",
]

# -- Analysis -----------------------------------------------------------------

a = Analysis(
    [str(PROJECT_ROOT / "launcher.py")],

    pathex=[str(PROJECT_ROOT)],

    binaries=[],

    datas=datas,

    hiddenimports=hidden_imports,

    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],

    excludes=[
        "tkinter",
        "matplotlib",
        "scipy",
        "IPython",
        "jupyter",
        "notebook",
        "pytest",
        "test",
        "tests",
    ],

    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

# -- PYZ archive --------------------------------------------------------------

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

# -- EXE ----------------------------------------------------------------------

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,

    name="StationDeck",

    debug=False,
    console=False,          # RELEASE: terminal hidden from end users

    bootloader_ignore_signals=False,
    strip=False,
    upx=False,

    icon=str(PROJECT_ROOT / "web" / "static" / "favicon.ico"),
)

# -- COLLECT ------------------------------------------------------------------

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="StationDeck",
)