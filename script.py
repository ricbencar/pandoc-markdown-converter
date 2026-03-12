# =============================================================================
# =============================================================================
# Pandoc Markdown Converter
# =============================================================================
#
# OBJECTIVE
# ---------
# This script converts a Markdown document into:
#   - DOCX, preserving mathematical expressions through Pandoc's native Office
#     Math conversion pipeline whenever the input uses Pandoc-compatible math.
#   - PDF, rendering formulas through Pandoc + LaTeX with a math-oriented
#     preamble designed to improve symbol coverage, Unicode handling, and AMS
#     environments.
#
# PDF OPENING PREFERENCES
# -----------------------
# The generated PDF is post-processed so it opens with:
#   - A4 paper format
#   - no bookmarks/thumbnails side panels
#   - one-column continuous reading mode
#   - fit-to-width initial view
#
# FORMULA RENDERING STRATEGY
# --------------------------
# For PDF, Pandoc does NOT use MathJax. MathJax is for HTML rendering. PDF math
# is compiled through a LaTeX engine. For that reason this script:
#
#   1) Prefers XeLaTeX, then LuaLaTeX, then pdfLaTeX:
#        --pdf-engine=xelatex
#        --pdf-engine=lualatex
#        --pdf-engine=pdflatex
#
#   2) Uses Pandoc Markdown with explicit math support:
#        -f markdown+tex_math_dollars+raw_tex
#
#   3) Injects a LaTeX header that loads robust math packages:
#        \usepackage{amsmath, amssymb, mathtools, bm}
#        \usepackage{unicode-math}   (for XeLaTeX/LuaLaTeX)
#
#   4) Normalizes common LaTeX display environments such as:
#        \begin{align} ... \end{align}
#        \begin{equation} ... \end{equation}
#        \begin{gather} ... \end{gather}
#        \[ ... \]
#
#      into Pandoc-friendly display math forms so DOCX conversion is more
#      reliable and PDF output remains stable.
#
#   5) Allows raw LaTeX to pass through for PDF output where needed:
#        markdown+raw_tex
#
# PYTHON PACKAGE INSTALLATION
# ---------------------------
# Install the required Python packages with pip:
#
#     pip install pypandoc pypdf
#
# EXTERNAL TOOLS REQUIRED
# -----------------------
# This script also requires tools that are NOT installed with pip:
#
#   1) Pandoc
#   2) A LaTeX PDF engine: xelatex, lualatex, or pdflatex
#
# The script automatically selects the first available engine in this order:
#   - xelatex
#   - lualatex
#   - pdflatex
#
# COMMON WINDOWS CUSTOM TEMPLATE LOCATION
# ---------------------------------------
# Custom Pandoc templates are usually placed in:
#
#     %APPDATA%\pandoc\templates
#
# The script also checks:
#   - Pandoc user data directory reported by `pandoc --version`
#   - %PANDOC_DATA_DIR%\templates
#   - ~/.pandoc/templates
#
# GUI USAGE
# ---------
# 1) Start the script with no arguments:
#
#        python script.py
#
# 2) Select the Markdown input file.
# 3) Review or edit the automatically generated DOCX/PDF output names.
# 4) Choose a template from the combo box.
# 5) Click Convert now.
#
# TEMPLATE NOTES
# --------------
# - Items marked "PDF OK" are compatible with the PDF route used in this script.
# - Items marked "PDF NO" can still be listed for reference, but the script will
#   automatically fall back to the default Pandoc PDF template.
# - If a LaTeX template references missing files such as a .cls or .sty file,
#   the script retries with the default Pandoc PDF template.
#
# CLI USAGE
# ---------
# Basic conversion:
#
#     python script.py README.md
#
# Explicit output names:
#
#     python script.py README.md --docx README.docx --pdf README.pdf
#
# Force GUI mode:
#
#     python script.py --gui
#
# List discoverable templates:
#
#     python script.py --list-templates
#
# Select a specific PDF template from CLI:
#
#     python script.py README.md --template latex
#
# OUTPUT NAMING
# -------------
# If the input is:
#
#     README.md
#
# then the default outputs are:
#
#     README.docx
#     README.pdf
#
# DEFAULT WORKING DIRECTORY
# -------------------------
# When the script starts, its working directory is changed to the directory
# where script.py is located. This ensures that the default input file,
# default output files, and file dialog starting locations are aligned with
# the script folder.
# =============================================================================

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

import pypandoc
from pypdf import PdfReader, PdfWriter
from pypdf.generic import ArrayObject, BooleanObject, FloatObject, NameObject

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk
except Exception:
    tk = None
    filedialog = None
    messagebox = None
    scrolledtext = None
    ttk = None


if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent

SCRIPT_DIR = APP_DIR
DEFAULT_INPUT = "README.md"

PANDOC_FORMAT = (
    "markdown+tex_math_dollars+raw_tex+raw_attribute+pipe_tables+"
    "fenced_code_blocks+backtick_code_blocks"
)

PDF_PAGE_LAYOUT = "/OneColumn"
PDF_PAGE_MODE = "/UseNone"
PDF_COMPATIBLE_TEMPLATE_FORMATS = {"latex"}

LATEX_MATH_HEADER = r"""
\usepackage{iftex}
\ifPDFTeX
  \usepackage[T1]{fontenc}
  \usepackage[utf8]{inputenc}
  \usepackage{lmodern}
\else
  \usepackage{fontspec}
  \defaultfontfeatures{Ligatures=TeX,Scale=MatchLowercase}
  \setmainfont{Latin Modern Roman}
  \setsansfont{Latin Modern Sans}
  \setmonofont{Latin Modern Mono}
  \usepackage{unicode-math}
  \setmathfont{Latin Modern Math}
\fi
\usepackage{amsmath,amssymb,mathtools,bm}
\allowdisplaybreaks
""".strip()

INSTRUCTIONS_TEXT = r"""
Pandoc Markdown Converter
=========================

Objectives
----------
This script converts a Markdown document into:
- DOCX with equations preserved through Pandoc's document conversion pipeline.
- PDF with formulas rendered through Pandoc + LaTeX.

It also applies PDF viewer preferences so the generated PDF opens with:
- A4 paper format
- no bookmarks/thumbnails panels
- one-column continuous reading mode
- fit-to-width initial view

Formula rendering notes
-----------------------
For PDF, Pandoc does not use MathJax. PDF formulas are compiled through a LaTeX
engine. For better mathematical rendering this script:
- prefers XeLaTeX or LuaLaTeX
- enables Pandoc math parsing with tex_math_dollars
- allows raw LaTeX when needed
- injects a math-oriented LaTeX header with AMS packages and unicode-math
- normalizes common AMS math environments for better DOCX/PDF compatibility

Python package installation
---------------------------
Install the required Python packages with pip:

    pip install pypandoc pypdf

External tools required
-----------------------
This script also requires tools that are NOT installed with pip:

1) Pandoc
2) A LaTeX PDF engine: xelatex, lualatex, or pdflatex

The script automatically selects the first available engine in this order:
- xelatex
- lualatex
- pdflatex

Common Windows custom template location
---------------------------------------
Custom Pandoc templates are usually placed in:

    %APPDATA%\pandoc\templates

The script also checks:
- Pandoc user data directory reported by `pandoc --version`
- %PANDOC_DATA_DIR%\templates
- ~/.pandoc/templates

GUI usage
---------
1) Start the script with no arguments:

       python script.py

2) Select the Markdown input file.
3) Review or edit the automatically generated DOCX/PDF output names.
4) Choose a template from the combo box.
5) Click Convert now.

Template notes
--------------
- Items marked "PDF OK" are compatible with the PDF route used in this script.
- Items marked "PDF NO" can still be listed for reference, but the script will
  automatically fall back to the default Pandoc PDF template.
- If a LaTeX template references missing files such as a .cls or .sty file, the
  script retries with the default Pandoc PDF template.

CLI usage
---------
Basic conversion:

    python script.py README.md

Explicit output names:

    python script.py README.md --docx README.docx --pdf README.pdf

Force GUI mode:

    python script.py --gui

List discoverable templates:

    python script.py --list-templates

Select a specific PDF template from CLI:

    python script.py README.md --template latex

Output naming
-------------
If the input is:

    README.md

then the default outputs are:

    README.docx
    README.pdf
""".strip()


LogFunc = Optional[Callable[[str], None]]


def log_message(logger: LogFunc, message: str) -> None:
    if logger:
        logger(message)


_ORIGINAL_SUBPROCESS_POPEN = subprocess.Popen


def _build_hidden_startupinfo(existing: Optional[subprocess.STARTUPINFO] = None):
    if os.name != "nt":
        return existing

    startupinfo = existing if existing is not None else subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE
    return startupinfo


def _apply_windows_no_console_kwargs(kwargs: Dict) -> Dict:
    if os.name != "nt":
        return kwargs

    kwargs = dict(kwargs)
    kwargs["creationflags"] = int(kwargs.get("creationflags", 0)) | subprocess.CREATE_NO_WINDOW
    kwargs["startupinfo"] = _build_hidden_startupinfo(kwargs.get("startupinfo"))
    return kwargs


def _hidden_popen(*args, **kwargs):
    return _ORIGINAL_SUBPROCESS_POPEN(*args, **_apply_windows_no_console_kwargs(kwargs))


def install_windows_subprocess_suppression() -> None:
    if os.name == "nt" and subprocess.Popen is not _hidden_popen:
        subprocess.Popen = _hidden_popen


install_windows_subprocess_suppression()


def run_command(args: List[str]) -> str:
    """Run a subprocess command and return stdout as text without opening console windows."""
    try:
        completed = subprocess.run(
            args,
            check=True,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        return completed.stdout
    except FileNotFoundError as exc:
        command_name = args[0] if args else "Command"
        raise RuntimeError("{} was not found in PATH.".format(command_name)) from exc
    except subprocess.CalledProcessError as exc:
        stderr = (exc.stderr or "").strip()
        stdout = (exc.stdout or "").strip()
        detail = stderr or stdout or "Unknown subprocess error."
        raise RuntimeError(detail) from exc


def get_pandoc_path() -> str:
    pandoc_path = shutil.which("pandoc")
    if not pandoc_path:
        raise RuntimeError("Pandoc was not found. Install the Pandoc binary and try again.")
    return pandoc_path


def get_pandoc_user_data_dir() -> Optional[Path]:
    """Extract Pandoc user data directory from `pandoc --version`."""
    _ = get_pandoc_path()
    version_text = run_command(["pandoc", "--version"])
    for line in version_text.splitlines():
        if line.startswith("User data directory:"):
            data_dir = line.split(":", 1)[1].strip()
            if data_dir:
                return Path(data_dir)
    return None


def choose_pdf_engine() -> str:
    """Return the best available LaTeX engine for PDF generation."""
    for engine in ("xelatex", "lualatex", "pdflatex"):
        if shutil.which(engine):
            return engine
    raise RuntimeError(
        "No supported LaTeX PDF engine found. Install xelatex, lualatex, or pdflatex."
    )


def resolve_input_path(input_file: str) -> Path:
    raw = (input_file or "").strip() or DEFAULT_INPUT
    path = Path(raw)
    if not path.is_absolute():
        path = SCRIPT_DIR / path
    return path.resolve()


def ensure_input_exists(input_file: str) -> Path:
    path = resolve_input_path(input_file)
    if not path.exists():
        raise FileNotFoundError("Input file '{}' not found.".format(path))
    return path


def derive_default_output_paths(input_file: str) -> Tuple[str, str]:
    """Derive DOCX/PDF output paths from the input file name."""
    input_path = resolve_input_path(input_file)
    parent = input_path.parent
    stem = input_path.stem if input_path.suffix else input_path.name

    if not stem:
        stem = "output"

    return str(parent / (stem + ".docx")), str(parent / (stem + ".pdf"))


def resolve_output_paths(
    input_file: str,
    docx_file: Optional[str],
    pdf_file: Optional[str],
) -> Tuple[str, str]:
    auto_docx, auto_pdf = derive_default_output_paths(input_file)

    if docx_file:
        docx_path = Path(docx_file)
        if not docx_path.is_absolute():
            docx_path = SCRIPT_DIR / docx_path
    else:
        docx_path = Path(auto_docx)

    if pdf_file:
        pdf_path = Path(pdf_file)
        if not pdf_path.is_absolute():
            pdf_path = SCRIPT_DIR / pdf_path
    else:
        pdf_path = Path(auto_pdf)

    return str(docx_path.resolve()), str(pdf_path.resolve())


def normalize_math_block_body(body: str) -> str:
    """
    Clean common LaTeX commands that often reduce Pandoc/Office Math robustness
    when converting display environments to Pandoc-native math blocks.
    """
    body = body.strip()
    body = re.sub(r"\\label\{.*?\}", "", body)
    body = re.sub(r"\\tag\*?\{.*?\}", "", body)
    body = body.replace(r"\nonumber", "")
    body = body.replace(r"\notag", "")
    return body.strip()


def convert_bracket_display_math(content: str) -> str:
    """Convert [...] display math to $$ ... $$."""
    pattern = re.compile(r"\\\[(.*?)\\\]", flags=re.DOTALL)

    def repl(match: re.Match) -> str:
        body = normalize_math_block_body(match.group(1))
        return "\n$$\n{}\n$$\n".format(body)

    return pattern.sub(repl, content)


def convert_latex_display_environments(content: str) -> str:
    """
    Convert common AMS display environments to Pandoc-friendly $$...$$ forms.

    This improves DOCX math conversion because raw LaTeX environments are often
    ignored outside LaTeX/PDF outputs, while Pandoc-native math blocks are
    converted to Office Math more reliably.
    """
    pattern = re.compile(
        r"(?ms)^[ \t]*\\begin\{(?P<env>equation\*?|align\*?|aligned|split|gather\*?|gathered|multline\*?|eqnarray\*?)\}\s*"
        r"(?P<body>.*?)"
        r"^[ \t]*\\end\{(?P=env)\}[ \t]*$"
    )

    def repl(match: re.Match) -> str:
        env = match.group("env")
        env_base = env.replace("*", "")
        body = normalize_math_block_body(match.group("body"))

        if env_base in {"align", "aligned", "split", "eqnarray"}:
            wrapped = "\\begin{aligned}\n" + body + "\n\\end{aligned}"
        elif env_base in {"gather", "gathered"}:
            wrapped = "\\begin{gathered}\n" + body + "\n\\end{gathered}"
        else:
            wrapped = body

        return "\n$$\n{}\n$$\n".format(wrapped)

    return pattern.sub(repl, content)


def preprocess_markdown(content: str) -> str:
    """
    Normalize Markdown so Pandoc reliably detects inline and display math.

    The preprocessing focuses on:
    - normalizing line endings
    - isolating $$ display math blocks
    - converting [...] to $$...$$
    - converting common AMS display environments to Pandoc-friendly math blocks
    - fixing a few problematic math constructs
    """
    content = content.replace("\r\n", "\n").replace("\r", "\n")
    content = convert_bracket_display_math(content)
    content = convert_latex_display_environments(content)

    # Ensure $$ display equations are isolated on their own lines.
    content = re.sub(r"([^\n])\$\$", r"\1\n$$", content)
    content = re.sub(r"\$\$([^\n])", r"$$\n\1", content)

    # Collapse excessive blank lines around display math blocks.
    content = re.sub(r"\n{3,}\$\$", r"\n\n$$", content)
    content = re.sub(r"\$\$\n{3,}", r"$$\n\n", content)

    return content


def create_temp_latex_header() -> str:
    """Create a temporary LaTeX header file with robust math configuration."""
    temp = tempfile.NamedTemporaryFile(
        mode="w",
        suffix=".tex",
        delete=False,
        encoding="utf-8",
    )
    try:
        temp.write(LATEX_MATH_HEADER + "\n")
        return temp.name
    finally:
        temp.close()


def apply_pdf_view_settings(pdf_file: str) -> None:
    """
    Apply PDF opening preferences:
    - no side panels
    - one-column continuous layout
    - fit to page width on first page
    """
    reader = PdfReader(pdf_file)
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    writer.page_mode = PDF_PAGE_MODE
    writer.set_page_layout(PDF_PAGE_LAYOUT)

    prefs = writer.create_viewer_preferences()
    # Do NOT set /FitWindow, because that means "resize window to page"
    # and is different from fit-to-width.
    prefs[NameObject("/CenterWindow")] = BooleanObject(True)
    prefs[NameObject("/DisplayDocTitle")] = BooleanObject(True)
    prefs[NameObject("/NonFullScreenPageMode")] = NameObject("/UseNone")

    if writer.pages:
        first_page = writer.pages[0]
        first_page_ref = first_page.indirect_reference
        if first_page_ref is not None:
            top_y = FloatObject(float(first_page.mediabox.top))
            writer._root_object.update({
                NameObject("/OpenAction"): ArrayObject([
                    first_page_ref,
                    NameObject("/FitH"),
                    top_y,
                ])
            })

    if reader.metadata:
        clean_metadata = {}
        for k, v in reader.metadata.items():
            if isinstance(k, str) and isinstance(v, str):
                clean_metadata[k] = v
        if clean_metadata:
            writer.add_metadata(clean_metadata)

    with open(pdf_file, "wb") as handle:
        writer.write(handle)


def _try_get_builtin_template(format_name: str) -> Optional[str]:
    try:
        return run_command(["pandoc", "-D", format_name])
    except Exception:
        return None


def discover_builtin_templates() -> List[Dict]:
    """Discover built-in Pandoc templates by probing writer formats."""
    _ = get_pandoc_path()
    formats_text = run_command(["pandoc", "--list-output-formats"])
    templates: List[Dict] = []

    for format_name in sorted(set(formats_text.split())):
        template_text = _try_get_builtin_template(format_name)
        if template_text is None:
            continue

        suffix = "latex" if format_name == "latex" else format_name
        pdf_compatible = format_name in PDF_COMPATIBLE_TEMPLATE_FORMATS

        # Keep only PDF-compatible built-in templates
        if not pdf_compatible:
            continue

        templates.append({
            "label": Path("default.{}".format(suffix)).stem,   # shown name only
            "source": "builtin",
            "format": format_name,
            "name": "default.{}".format(suffix),
            "path": None,
            "content": template_text,
            "pdf_compatible": True,
        })

    return templates


def discover_custom_templates() -> List[Dict]:
    """Discover custom templates in standard Pandoc template directories."""
    search_dirs: List[Path] = []

    user_data_dir = get_pandoc_user_data_dir()
    if user_data_dir:
        search_dirs.append(user_data_dir / "templates")

    env_data_dir = os.environ.get("PANDOC_DATA_DIR", "").strip()
    if env_data_dir:
        search_dirs.append(Path(env_data_dir) / "templates")

    appdata_dir = os.environ.get("APPDATA", "").strip()
    if appdata_dir:
        search_dirs.append(Path(appdata_dir) / "pandoc" / "templates")

    legacy_dir = Path.home() / ".pandoc" / "templates"
    search_dirs.append(legacy_dir)

    templates: List[Dict] = []
    seen = set()

    for directory in search_dirs:
        if not directory.exists() or not directory.is_dir():
            continue

        for path in sorted(directory.rglob("*")):
            if not path.is_file():
                continue

            key = str(path.resolve()).lower()
            if key in seen:
                continue
            seen.add(key)

            suffix = path.suffix.lower().lstrip(".")
            pdf_compatible = suffix == "latex"
            compat_tag = "PDF OK" if pdf_compatible else "PDF NO"
            label = "custom | {} | {} | {} | {}".format(
                suffix or "file",
                path.name,
                str(path),
                compat_tag,
            )

            templates.append({
                "label": label,
                "source": "custom",
                "format": suffix or "unknown",
                "name": path.name,
                "path": str(path),
                "content": None,
                "pdf_compatible": pdf_compatible,
            })

    return templates


def discover_pandoc_templates() -> List[Dict]:
    def clean_name(item: Dict) -> str:
        # Default Pandoc PDF template
        if item.get("source") == "default":
            return "default"

        # Built-in template
        if item.get("source") == "builtin":
            name = item.get("name", "") or "default"
            stem = Path(name).stem
            return stem or "default"

        # Custom template file
        path_value = item.get("path")
        if path_value:
            return Path(path_value).stem

        name = item.get("name", "") or "template"
        return Path(name).stem or "template"

    template_items = [{
        "label": "default",
        "source": "default",
        "format": "latex",
        "name": "(default)",
        "path": None,
        "content": None,
        "pdf_compatible": True,
    }]

    template_items.extend(discover_builtin_templates())
    template_items.extend(discover_custom_templates())

    # Keep only PDF-compatible templates
    template_items = [item for item in template_items if item.get("pdf_compatible")]

    # Replace verbose labels with clean template names
    for item in template_items:
        item["label"] = clean_name(item)

    # If two templates have the same visible name, disambiguate with a numeric suffix
    seen = {}
    for item in template_items:
        base = item["label"]
        count = seen.get(base, 0) + 1
        seen[base] = count
        if count > 1:
            item["label"] = "{} ({})".format(base, count)

    template_items.sort(key=lambda item: item["label"].lower())
    return template_items


def export_builtin_template_to_tempfile(template_info: Dict) -> str:
    format_name = template_info["format"]
    suffix = ".latex" if format_name == "latex" else ".{}".format(format_name)
    temp = tempfile.NamedTemporaryFile(
        mode="w",
        suffix=suffix,
        delete=False,
        encoding="utf-8",
    )
    try:
        temp.write(template_info.get("content") or "")
        return temp.name
    finally:
        temp.close()


def resolve_template_info(template_value: Optional[str]) -> Optional[Dict]:
    if not template_value:
        return None

    path = Path(template_value)
    if path.exists() and path.is_file():
        return {
            "label": "custom | {} | {} | {}".format(
                path.suffix.lstrip(".") or "file",
                path.name,
                str(path),
            ),
            "source": "custom",
            "format": path.suffix.lower().lstrip(".") or "unknown",
            "name": path.name,
            "path": str(path),
            "content": None,
            "pdf_compatible": path.suffix.lower() == ".latex",
        }

    builtin_template = _try_get_builtin_template(template_value)
    if builtin_template is not None:
        return {
            "label": "builtin | {} | default.{}".format(template_value, template_value),
            "source": "builtin",
            "format": template_value,
            "name": "default.{}".format(template_value),
            "path": None,
            "content": builtin_template,
            "pdf_compatible": template_value in PDF_COMPATIBLE_TEMPLATE_FORMATS,
        }

    raise ValueError(
        "Invalid template selection. Use a template file path or a built-in format such as `latex`."
    )


def build_pdf_template_args(
    template_info: Optional[Dict],
) -> Tuple[List[str], List[str], Optional[str]]:
    """
    Build Pandoc `--template` arguments for PDF conversion.

    Non-LaTeX templates do not abort the conversion. The function returns a warning
    and the caller falls back to the default Pandoc PDF template.
    """
    if not template_info or template_info.get("source") == "default":
        return [], [], None

    if not template_info.get("pdf_compatible", False):
        warning = (
            "The selected template is not compatible with PDF generation in this script. "
            "The default Pandoc PDF template will be used instead."
        )
        return [], [], warning

    cleanup_files: List[str] = []

    if template_info["source"] == "builtin":
        temp_template = export_builtin_template_to_tempfile(template_info)
        cleanup_files.append(temp_template)
        return ["--template={}".format(temp_template)], cleanup_files, None

    if template_info["source"] == "custom":
        return ["--template={}".format(template_info["path"])], cleanup_files, None

    return [], [], None


def needs_latex_fallback(error_text: str) -> bool:
    lower = error_text.lower()
    return (
        ".cls not found" in lower
        or ".sty not found" in lower
        or "latex error: file " in lower
        or ("error producing pdf" in lower and ("type x to quit" in lower or "emergency stop" in lower))
    )


def convert_markdown_file(
    input_file: str,
    docx_file: Optional[str],
    pdf_file: Optional[str],
    template_info: Optional[Dict] = None,
    logger: LogFunc = None,
) -> Tuple[str, str, Optional[str]]:
    """
    Convert Markdown to DOCX and PDF.

    DOCX:
    - uses Pandoc's native DOCX writer, which converts Pandoc-compatible TeX math
      to Office Math

    PDF:
    - uses Pandoc + LaTeX with a math-focused include header
    - retries with the default template if a selected template depends on missing
      LaTeX class/style files
    """
    input_path = ensure_input_exists(input_file)
    final_docx_file, final_pdf_file = resolve_output_paths(str(input_path), docx_file, pdf_file)

    with open(input_path, "r", encoding="utf-8") as handle:
        original_content = handle.read()

    content = preprocess_markdown(original_content)
    log_message(logger, "Read {} characters from {}.".format(len(original_content), input_path))

    pdf_engine = choose_pdf_engine()
    log_message(logger, "Using PDF engine: {}.".format(pdf_engine))

    temp_md_path = None
    temp_header_path = None
    template_cleanup_files: List[str] = []
    template_warning: Optional[str] = None

    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            suffix=".md",
            delete=False,
            encoding="utf-8",
        ) as temp_md:
            temp_md.write(content)
            temp_md_path = temp_md.name

        temp_header_path = create_temp_latex_header()
        common_args = ["--standalone"]

        log_message(logger, "Converting to DOCX: {}.".format(final_docx_file))
        pypandoc.convert_file(
            temp_md_path,
            to="docx",
            format=PANDOC_FORMAT,
            outputfile=final_docx_file,
            extra_args=common_args,
        )

        pdf_template_args, template_cleanup_files, template_warning = build_pdf_template_args(template_info)
        pdf_args = common_args + pdf_template_args + [
            "--pdf-engine={}".format(pdf_engine),
            "--include-in-header={}".format(temp_header_path),
            "-V", "papersize:a4",
            "-V", "geometry:margin=1in",
            "-V", "fontsize=11pt",
            "-V", "linestretch=1.08",
            "-V", "colorlinks=true",
        ]

        log_message(logger, "Converting to PDF: {}.".format(final_pdf_file))
        try:
            pypandoc.convert_file(
                temp_md_path,
                to="pdf",
                format=PANDOC_FORMAT,
                outputfile=final_pdf_file,
                extra_args=pdf_args,
            )
        except RuntimeError as exc:
            err = str(exc)
            selected_custom_template = bool(template_info and template_info.get("source") != "default")
            if selected_custom_template and needs_latex_fallback(err):
                fallback_note = (
                    "The selected LaTeX template requires external LaTeX files that are not "
                    "installed. The default Pandoc PDF template was used instead."
                )
                template_warning = (
                    "{} {}".format(template_warning, fallback_note).strip()
                    if template_warning else fallback_note
                )
                log_message(logger, "Selected template failed. Retrying PDF with default Pandoc template.")

                for path in template_cleanup_files:
                    if path and os.path.exists(path):
                        os.remove(path)
                template_cleanup_files = []

                fallback_args = common_args + [
                    "--pdf-engine={}".format(pdf_engine),
                    "--include-in-header={}".format(temp_header_path),
                    "-V", "papersize:a4",
                    "-V", "geometry:margin=1in",
                    "-V", "fontsize=11pt",
                    "-V", "linestretch=1.08",
                    "-V", "colorlinks=true",
                ]
                pypandoc.convert_file(
                    temp_md_path,
                    to="pdf",
                    format=PANDOC_FORMAT,
                    outputfile=final_pdf_file,
                    extra_args=fallback_args,
                )
            else:
                raise

        apply_pdf_view_settings(final_pdf_file)
        log_message(logger, "PDF viewer preferences applied successfully.")

    except OSError as exc:
        raise RuntimeError(
            "Pandoc was not found. Install the Pandoc binary and try again."
        ) from exc
    finally:
        if temp_md_path and os.path.exists(temp_md_path):
            os.remove(temp_md_path)
        if temp_header_path and os.path.exists(temp_header_path):
            os.remove(temp_header_path)
        for path in template_cleanup_files:
            if path and os.path.exists(path):
                os.remove(path)

    return os.path.abspath(final_docx_file), os.path.abspath(final_pdf_file), template_warning


class PandocConverterGUI:
    """Tk interface for the Markdown converter."""

    def __init__(self, root: "tk.Tk"):
        self.root = root
        self.root.title("Pandoc Markdown Converter")
        self.root.geometry("900x680")
        self.root.minsize(820, 620)

        self.input_var = tk.StringVar(value=DEFAULT_INPUT)
        self.docx_var = tk.StringVar(value="")
        self.pdf_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(
            value="Ready. Configure the conversion in the first tab and run it."
        )
        self.template_summary_var = tk.StringVar(value="Template list not loaded yet.")

        self.template_items: List[Dict] = []
        self.template_map: Dict[str, Dict] = {}
        self._last_auto_docx = ""
        self._last_auto_pdf = ""

        self._configure_style()
        self._build_header()
        self._build_body()
        self._build_footer()
        self._auto_update_output_names(force=True)
        self.input_var.trace_add("write", self._on_input_changed)
        self.refresh_templates()

    def _configure_style(self) -> None:
        theme_candidates = ("vista", "xpnative", "clam", "alt", "default")
        style = ttk.Style(self.root)
        for theme_name in theme_candidates:
            if theme_name in style.theme_names():
                try:
                    style.theme_use(theme_name)
                    break
                except Exception:
                    pass

        self.root.configure(bg="#eef3f8")
        style.configure("TNotebook", background="#eef3f8", borderwidth=0)
        style.configure("TNotebook.Tab", padding=(16, 8), font=("Segoe UI", 10, "bold"))
        style.configure("Card.TLabelframe", padding=10)
        style.configure("Card.TLabelframe.Label", font=("Segoe UI", 10, "bold"))
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("TButton", padding=(10, 6), font=("Segoe UI", 10))
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("Small.TLabel", font=("Segoe UI", 9))
        style.configure("Status.TLabel", font=("Segoe UI", 10))
        style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"))
        style.configure("Subtitle.TLabel", font=("Segoe UI", 10))

    def _build_header(self) -> None:
        header = tk.Frame(self.root, bg="#1f3b5b", padx=16, pady=14)
        header.pack(fill="x")

        title = tk.Label(
            header,
            text="Pandoc Markdown Converter",
            bg="#1f3b5b",
            fg="white",
            font=("Segoe UI", 18, "bold"),
            anchor="w",
        )
        title.pack(fill="x")

        subtitle = tk.Label(
            header,
            text=(
                "DOCX + PDF generation with Pandoc, LaTeX-rendered formulas, "
                "template discovery, and PDF opening preferences."
            ),
            bg="#1f3b5b",
            fg="#d7e6f5",
            font=("Segoe UI", 10),
            anchor="w",
        )
        subtitle.pack(fill="x", pady=(4, 0))

    def _build_body(self) -> None:
        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill="both", expand=True)

        self.notebook = ttk.Notebook(outer)
        self.notebook.pack(fill="both", expand=True)

        self.convert_tab = ttk.Frame(self.notebook, padding=14)
        self.log_tab = ttk.Frame(self.notebook, padding=14)
        self.instructions_tab = ttk.Frame(self.notebook, padding=14)

        self.notebook.add(self.convert_tab, text="Convert")
        self.notebook.add(self.log_tab, text="Log")
        self.notebook.add(self.instructions_tab, text="Instructions")

        self._build_convert_tab()
        self._build_log_tab()
        self._build_instructions_tab()

        self.notebook.select(self.convert_tab)

    def _build_instructions_tab(self) -> None:
        info_frame = ttk.LabelFrame(
            self.instructions_tab,
            text="Script objectives and detailed usage",
            style="Card.TLabelframe",
        )
        info_frame.pack(fill="both", expand=True)

        overview_box = scrolledtext.ScrolledText(
            info_frame,
            wrap="word",
            font=("Consolas", 12),
            padx=12,
            pady=12,
            borderwidth=0,
            relief="flat",
            background="white",
        )
        overview_box.pack(fill="both", expand=True)
        overview_box.insert("1.0", INSTRUCTIONS_TEXT)
        overview_box.configure(state="disabled")

    def _build_convert_tab(self) -> None:
        self.convert_tab.columnconfigure(0, weight=3)
        self.convert_tab.columnconfigure(1, weight=2)
        self.convert_tab.rowconfigure(0, weight=1)

        left = ttk.Frame(self.convert_tab)
        right = ttk.Frame(self.convert_tab)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        right.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        left.columnconfigure(0, weight=1)
        right.columnconfigure(0, weight=1)

        files_card = ttk.LabelFrame(left, text="Files", style="Card.TLabelframe")
        files_card.grid(row=0, column=0, sticky="ew")
        files_card.columnconfigure(1, weight=1)

        ttk.Label(files_card, text="Input Markdown").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=6)
        ttk.Entry(files_card, textvariable=self.input_var).grid(row=0, column=1, sticky="ew", pady=6)
        ttk.Button(files_card, text="Browse", command=self.browse_input).grid(row=0, column=2, sticky="ew", padx=(8, 0), pady=6)

        ttk.Label(files_card, text="Output DOCX").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=6)
        ttk.Entry(files_card, textvariable=self.docx_var).grid(row=1, column=1, sticky="ew", pady=6)
        ttk.Button(files_card, text="Browse", command=self.browse_docx).grid(row=1, column=2, sticky="ew", padx=(8, 0), pady=6)

        ttk.Label(files_card, text="Output PDF").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=6)
        ttk.Entry(files_card, textvariable=self.pdf_var).grid(row=2, column=1, sticky="ew", pady=6)
        ttk.Button(files_card, text="Browse", command=self.browse_pdf).grid(row=2, column=2, sticky="ew", padx=(8, 0), pady=6)

        notes = ttk.Label(
            files_card,
            text="Output names track the input file name until you manually replace them.",
            style="Small.TLabel",
            wraplength=520,
            justify="left",
        )
        notes.grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 2))

        template_card = ttk.LabelFrame(left, text="Pandoc templates", style="Card.TLabelframe")
        template_card.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        template_card.columnconfigure(0, weight=1)

        self.lcombobox = ttk.Combobox(template_card, state="readonly")
        self.lcombobox.grid(row=0, column=0, sticky="ew", pady=6)
        self.lcombobox.bind("<<ComboboxSelected>>", self._on_template_changed)
        ttk.Button(template_card, text="Refresh templates", command=self.refresh_templates).grid(
            row=0, column=1, sticky="ew", padx=(8, 0), pady=6
        )

        ttk.Label(
            template_card,
            textvariable=self.template_summary_var,
            wraplength=520,
            justify="left",
            style="Small.TLabel",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        action_card = ttk.LabelFrame(right, text="Actions", style="Card.TLabelframe")
        action_card.grid(row=0, column=0, sticky="new")
        action_card.columnconfigure(0, weight=1)

        ttk.Button(action_card, text="Convert now", style="Primary.TButton", command=self.convert).grid(
            row=0, column=0, sticky="ew", pady=(0, 8)
        )
        ttk.Button(action_card, text="Open Log", command=lambda: self.notebook.select(self.log_tab)).grid(
            row=1, column=0, sticky="ew", pady=4
        )
        ttk.Button(action_card, text="Open Instructions", command=lambda: self.notebook.select(self.instructions_tab)).grid(
            row=2, column=0, sticky="ew", pady=4
        )
        ttk.Button(action_card, text="Quit", command=self.root.destroy).grid(
            row=3, column=0, sticky="ew", pady=4
        )

        summary_card = ttk.LabelFrame(right, text="Current conversion profile", style="Card.TLabelframe")
        summary_card.grid(row=1, column=0, sticky="new", pady=(12, 0))

        summary_text = (
            "PDF settings:\n"
            "- paper size: A4\n"
            "- initial page mode: no side panels\n"
            "- page layout: one-column continuous\n"
            "- open action: fit to width\n\n"
            "Math rendering:\n"
            "- Pandoc + LaTeX\n"
            "- engine priority: xelatex, lualatex, pdflatex\n"
            "- LaTeX math header: amsmath, amssymb, mathtools, bm, unicode-math"
        )
        ttk.Label(summary_card, text=summary_text, justify="left").pack(fill="x")

    def _build_log_tab(self) -> None:
        log_frame = ttk.LabelFrame(self.log_tab, text="Execution log", style="Card.TLabelframe")
        log_frame.pack(fill="both", expand=True)

        self.log_box = scrolledtext.ScrolledText(
            log_frame,
            wrap="word",
            font=("Consolas", 12),
            padx=12,
            pady=12,
            borderwidth=0,
            relief="flat",
            background="white",
        )
        self.log_box.pack(fill="both", expand=True)
        self.log_box.configure(state="disabled")

    def _build_footer(self) -> None:
        footer = ttk.Frame(self.root, padding=(12, 0, 12, 12))
        footer.pack(fill="x")
        ttk.Label(
            footer,
            textvariable=self.status_var,
            style="Status.TLabel",
            wraplength=860,
            justify="left",
        ).pack(fill="x")

    def _append_log(self, message: str) -> None:
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message.rstrip() + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.root.update_idletasks()

    def _on_input_changed(self, *args: object) -> None:
        self._auto_update_output_names(force=False)

    def _auto_update_output_names(self, force: bool = False) -> None:
        auto_docx, auto_pdf = derive_default_output_paths(self.input_var.get().strip())
        current_docx = self.docx_var.get().strip()
        current_pdf = self.pdf_var.get().strip()

        if force or not current_docx or current_docx == self._last_auto_docx:
            self.docx_var.set(auto_docx)
        if force or not current_pdf or current_pdf == self._last_auto_pdf:
            self.pdf_var.set(auto_pdf)

        self._last_auto_docx = auto_docx
        self._last_auto_pdf = auto_pdf

    def _on_template_changed(self, event: object = None) -> None:
        label = self.lcombobox.get().strip()
        item = self.template_map.get(label)
        if not item:
            self.template_summary_var.set("No template selected.")
            return

        source = item.get("source", "unknown")
        fmt = item.get("format", "unknown")
        compat = "PDF OK" if item.get("pdf_compatible") else "PDF NO"
        extra = item.get("path") or item.get("name") or "(default)"
        self.template_summary_var.set(
            "Source: {} | Format: {} | {} | {}".format(source, fmt, compat, extra)
        )

    def browse_input(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Markdown file",
            initialdir=str(SCRIPT_DIR),
            filetypes=[("Markdown files", "*.md *.markdown *.txt"), ("All files", "*.*")],
        )
        if path:
            self.input_var.set(path)

    def browse_docx(self) -> None:
        current_name = Path(self.docx_var.get() or (SCRIPT_DIR / "output.docx")).name
        path = filedialog.asksaveasfilename(
            title="Save DOCX as",
            initialdir=str(SCRIPT_DIR),
            initialfile=current_name,
            defaultextension=".docx",
            filetypes=[("Word document", "*.docx")],
        )
        if path:
            self.docx_var.set(path)

    def browse_pdf(self) -> None:
        current_name = Path(self.pdf_var.get() or (SCRIPT_DIR / "output.pdf")).name
        path = filedialog.asksaveasfilename(
            title="Save PDF as",
            initialdir=str(SCRIPT_DIR),
            initialfile=current_name,
            defaultextension=".pdf",
            filetypes=[("PDF file", "*.pdf")],
        )
        if path:
            self.pdf_var.set(path)

    def refresh_templates(self) -> None:
        try:
            self.template_items = discover_pandoc_templates()
            self.template_map = {item["label"]: item for item in self.template_items}
            values = [item["label"] for item in self.template_items]
            self.lcombobox["values"] = values

            if values:
                selected = None

                # Prefer the Pandoc default template explicitly
                for item in self.template_items:
                    name = str(item.get("name", "")).strip().lower()
                    label = str(item.get("label", "")).strip().lower()
                    source = str(item.get("source", "")).strip().lower()

                    if (
                        name == "default"
                        or name.startswith("default.")
                        or label == "default"
                        or label.startswith("default ")
                        or source == "default"
                    ):
                        selected = item["label"]
                        break

                # Fallback: first PDF-compatible template
                if selected is None:
                    for item in self.template_items:
                        if item.get("pdf_compatible"):
                            selected = item["label"]
                            break

                # Final fallback: first item in the list
                if selected is None:
                    selected = values[0]

                self.lcombobox.set(selected)
                self._on_template_changed()

            self.status_var.set("Discovered {} Pandoc template options.".format(len(values)))
            self._append_log("Template refresh complete. {} options found.".format(len(values)))
        except Exception as exc:
            self.template_items = []
            self.template_map = {}
            self.lcombobox["values"] = []
            self.template_summary_var.set("Failed to load templates.")
            self.status_var.set("Failed to list templates: {}".format(exc))
            self._append_log("Template refresh failed: {}".format(exc))

    def convert(self) -> None:
        label = self.lcombobox.get().strip()
        template_info = self.template_map.get(label)

        self._append_log("--- Conversion started ---")
        self.notebook.select(self.log_tab)

        try:
            docx_path, pdf_path, template_warning = convert_markdown_file(
                self.input_var.get().strip(),
                self.docx_var.get().strip(),
                self.pdf_var.get().strip(),
                template_info=template_info,
                logger=self._append_log,
            )

            status = "Success. DOCX: {} | PDF: {}".format(docx_path, pdf_path)
            if template_warning:
                status += " | Note: {}".format(template_warning)
            self.status_var.set(status)
            self._append_log("Conversion completed successfully.")

            message = "DOCX:\n{}\n\nPDF:\n{}".format(docx_path, pdf_path)
            if template_warning:
                message += "\n\nNote:\n{}".format(template_warning)
            # messagebox.showinfo("Conversion complete", message)
        except Exception as exc:
            self.status_var.set("Error: {}".format(exc))
            self._append_log("Conversion failed: {}".format(exc))
            # messagebox.showerror("Conversion failed", str(exc))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Convert a Markdown file to DOCX and PDF with properly rendered formulas. "
            "Default mode is GUI when no command-line arguments are supplied."
        )
    )
    parser.add_argument(
        "input",
        nargs="?",
        default=DEFAULT_INPUT,
        help="Input Markdown file (default: {}).".format(DEFAULT_INPUT),
    )
    parser.add_argument(
        "--docx",
        default=None,
        help="Output DOCX path. If omitted, it follows the input file name.",
    )
    parser.add_argument(
        "--pdf",
        default=None,
        help="Output PDF path. If omitted, it follows the input file name.",
    )
    parser.add_argument(
        "--template",
        default=None,
        help="PDF template selector: a `.latex` template path or a built-in pandoc format such as `latex`.",
    )
    parser.add_argument(
        "--list-templates",
        action="store_true",
        help="List discoverable Pandoc templates and exit.",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Force GUI mode.",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    os.chdir(SCRIPT_DIR)

    if args.list_templates:
        try:
            for item in discover_pandoc_templates():
                print(item["label"])
        except Exception as exc:
            print("ERROR: {}".format(exc))
            return 1
        return 0

    no_cli_args = len(sys.argv) == 1
    if args.gui or no_cli_args:
        if tk is None:
            print("ERROR: Tkinter is not available in this Python installation.")
            return 1
        root = tk.Tk()
        PandocConverterGUI(root)
        root.mainloop()
        return 0

    try:
        template_info = resolve_template_info(args.template)
        docx_path, pdf_path, template_warning = convert_markdown_file(
            args.input,
            args.docx,
            args.pdf,
            template_info=template_info,
        )
    except Exception as exc:
        print("ERROR: {}".format(exc))
        return 1

    print("Success! Generated DOCX: {}".format(docx_path))
    print("Success! Generated PDF : {}".format(pdf_path))
    if template_warning:
        print("Note: {}".format(template_warning))
    print("PDF configured for A4, no side panels, continuous one-column view, and Fit-to-width opening.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())