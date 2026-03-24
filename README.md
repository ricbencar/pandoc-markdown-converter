# Pandoc Markdown Converter

A single-file Python application that converts Markdown documents into **DOCX** and **PDF** using **Pandoc**, with particular attention to **mathematical content**, **LaTeX compatibility**, **template discovery**, and **PDF opening behaviour**.

This project is designed for users who need a practical conversion workflow for technical or scientific Markdown files, especially when formulas must remain usable in both:

- **Microsoft Word (`.docx`)**, through Pandoc's Office Math conversion pipeline
- **PDF (`.pdf`)**, through Pandoc + LaTeX rendering

The script supports both a **graphical user interface (GUI)** and a **command-line interface (CLI)**. When started with no arguments, it opens the GUI by default.

---

## Contents

- [Overview](#overview)
- [Main capabilities](#main-capabilities)
- [How the conversion pipeline works](#how-the-conversion-pipeline-works)
- [Math handling strategy](#math-handling-strategy)
- [PDF behaviour and viewer settings](#pdf-behaviour-and-viewer-settings)
- [Template discovery and selection](#template-discovery-and-selection)
- [Requirements](#requirements)
- [Installation](#installation)
- [Standalone executable build with PyInstaller](#standalone-executable-build-with-pyinstaller)
- [Quick start](#quick-start)
- [GUI usage](#gui-usage)
- [CLI usage](#cli-usage)
- [Command-line arguments](#command-line-arguments)
- [Examples](#examples)
- [Project structure](#project-structure)
- [Key implementation details](#key-implementation-details)
- [Error handling and fallbacks](#error-handling-and-fallbacks)
- [Troubleshooting](#troubleshooting)
- [Limitations](#limitations)
- [Recommended use cases](#recommended-use-cases)

---

## Overview

`script.py` is a conversion utility built around Pandoc. Its purpose is to provide a more reliable Markdown-to-DOCX/PDF workflow than a minimal Pandoc wrapper, particularly for documents that contain:

- inline mathematics
- display equations
- common LaTeX math environments
- Unicode mathematical symbols
- technical formatting that benefits from a controlled PDF generation route

The application does more than simply call Pandoc:

1. It **preprocesses Markdown** to improve math robustness.
2. It **selects an available LaTeX engine** for PDF generation.
3. It **injects a LaTeX header** with math-oriented packages.
4. It **discovers available Pandoc templates**.
5. It **defaults to the standard Pandoc template** unless another valid template is selected.
6. It **post-processes the generated PDF** to enforce opening preferences such as:
   - A4 output intent in the Pandoc conversion stage
   - no side panels
   - one-column continuous layout
   - fit-to-width initial opening view

This makes the tool suitable for technical notes, reports, research drafts, engineering documentation, and formula-heavy Markdown content.

---

## Main capabilities

### Document conversion

The script converts one Markdown input file into:

- one **DOCX** file
- one **PDF** file

### GUI and CLI modes

The script supports two execution styles:

- **GUI mode** for interactive use
- **CLI mode** for scripted or terminal-based workflows

If you run the script with no arguments, the **GUI opens automatically**.

### Automatic output naming

If the input file is, for example:

```text
README.md
```

the default outputs become:

```text
README.docx
README.pdf
```

### Mathematical content support

The script is designed to improve handling of:

- `$...$` inline math
- `$$...$$` display math
- `\[ ... \]`
- `equation`, `equation*`
- `align`, `align*`
- `aligned`
- `split`
- `gather`, `gather*`
- `gathered`
- `multline`, `multline*`
- `eqnarray`, `eqnarray*`

### Template management

The application can:

- discover the default Pandoc PDF template route
- discover built-in Pandoc templates that are PDF-compatible
- discover custom templates from common Pandoc template directories
- populate a GUI dropdown with available template names
- default the GUI selection to **`default`**

### PDF post-processing

After PDF creation, the script modifies viewer preferences so the PDF opens with:

- **no bookmarks/thumbnails side panels**
- **one-column continuous layout**
- **fit-to-width initial view**

---

## How the conversion pipeline works

The conversion process is not a single direct call. It is a structured pipeline.

### 1. Input resolution

The script resolves the input path with the following behaviour:

- if no explicit input is provided, it uses `README.md`
- if the path is relative, it is interpreted relative to the script directory
- the working directory is changed to the folder that contains `script.py`

This ensures consistent behaviour for:

- input discovery
- default output naming
- GUI browse dialogs
- relative output paths

### 2. Markdown preprocessing

Before Pandoc is called, the Markdown content is normalised. This improves compatibility with both DOCX and PDF conversion.

The preprocessing stage includes:

- normalising line endings
- converting `\[ ... \]` to `$$ ... $$`
- converting supported LaTeX display environments into Pandoc-friendly display math blocks
- isolating `$$` display equations on their own lines
- reducing excessive blank lines around display math
- replacing a few specific math constructs that commonly cause downstream issues

### 3. DOCX generation

The DOCX file is produced using Pandoc's DOCX writer. The script relies on Pandoc's native ability to convert compatible TeX-like math into **Office Math**.

### 4. PDF generation

The PDF file is produced using:

- Pandoc
- a selected LaTeX engine
- an injected LaTeX header with math packages
- explicit PDF variables such as paper size and font size

### 5. PDF viewer preference update

After the PDF has been generated, the script reopens it with `pypdf` and writes viewer preferences into the file metadata structure.

---

## Math handling strategy

A central design goal of this project is to improve the reliability of formula conversion.

### Why the script preprocesses formulas

Pandoc generally handles Markdown math well, but raw LaTeX environments are not equally robust across all output formats.

In particular:

- **PDF generation** can handle many LaTeX math structures because LaTeX is part of the rendering route.
- **DOCX generation** is more reliable when math appears in Pandoc-friendly math forms rather than arbitrary raw LaTeX blocks.

For that reason, the script converts several display-math patterns into a more standardised representation before conversion.

### Conversions performed

#### Bracket display math

This form:

```latex
\[
E = mc^2
\]
```

is converted into:

```latex
$$
E = mc^2
$$
```

#### Common AMS environments

Supported environments are mapped into `$$ ... $$` blocks. Some environments are wrapped as aligned or gathered structures where appropriate.

Examples include:

- `align` -> wrapped into `aligned`
- `gather` -> wrapped into `gathered`
- `equation` -> simplified to a display block

### Cleanup of problematic commands

Inside converted math blocks, the script removes or neutralises some commands that often reduce conversion robustness, such as:

- `\label{...}`
- `\tag{...}`
- `\tag*{...}`
- `\nonumber`
- `\notag`

---

## PDF behaviour and viewer settings

The script does not stop at creating a PDF. It also writes viewer preferences so the generated file opens in a more controlled reading mode.

### PDF generation settings

During Pandoc PDF generation, the script passes variables equivalent to:

- `papersize:a4`
- `geometry:margin=1in`
- `fontsize=11pt`
- `linestretch=1.08`
- `colorlinks=true`

### Viewer preferences applied afterwards

After PDF creation, the script uses `pypdf` to apply the following opening preferences:

- `page_mode = /UseNone`
  - avoids opening bookmark or thumbnail side panels
- `page_layout = /OneColumn`
  - requests one-column continuous reading layout
- `CenterWindow = true`
- `DisplayDocTitle = true`
- `NonFullScreenPageMode = /UseNone`
- `OpenAction = FitH` on the first page

### Important note about viewers

PDF viewers vary in how strictly they honour embedded opening preferences. Many viewers respect these settings, but behaviour is not guaranteed to be identical across all software.

---

## Template discovery and selection

The script includes a template discovery subsystem focused on PDF-compatible templates.

### Default template

The application always includes an explicit **`default`** option. In the GUI, the selection logic is designed to prefer this option automatically.

### Built-in templates

The script probes Pandoc built-in templates by checking available output formats and calling:

```bash
pandoc -D <format>
```

Only **PDF-compatible built-in templates** are kept. In the current implementation, PDF compatibility is restricted to:

- `latex`

### Custom templates

The script also scans common template directories such as:

- Pandoc user data directory reported by `pandoc --version`
- `%PANDOC_DATA_DIR%\templates`
- `%APPDATA%\pandoc\templates`
- `~/.pandoc/templates`

Only templates with a `.latex` extension are treated as PDF-compatible.

### Display in the GUI

The template dropdown shows a cleaned label rather than a verbose metadata string. If duplicate visible names occur, the script disambiguates them with numeric suffixes.

### Fallback behaviour

If a selected custom or built-in LaTeX template fails because it depends on missing external LaTeX files such as a class or style file, the script retries PDF generation with the **default Pandoc template**.

---

## Requirements

### Python

The script requires Python 3 and imports at least these Python packages:

- `pypandoc`
- `pypdf`

The import is explicit in the script:

```python
import pypandoc
from pypdf import PdfReader, PdfWriter
```

### External software

The following tools must also be installed separately:

- **Pandoc**
- at least one LaTeX PDF engine:
  - `xelatex`
  - `lualatex`
  - `pdflatex`

### GUI support

For GUI mode, Python must include **Tkinter**.

---

## Installation

The recommended setup is to create a dedicated **virtual environment** for this project and install all Python dependencies into that environment.

### Windows: full setup from scratch

Open **Command Prompt** or **PowerShell** in the project folder and run:

```bat
cd C:\path\to\your\project
py -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip setuptools wheel
python -m pip install pypandoc pypdf
```

If you prefer PowerShell and script execution is blocked, you may need:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Then activate again:

```powershell
.\.venv\Scripts\Activate.ps1
```

### Linux or macOS: full setup from scratch

```bash
cd /path/to/your/project
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip setuptools wheel
python -m pip install pypandoc pypdf
```

### Why `python -m pip` is recommended

Using:

```bash
python -m pip install ...
```

is more robust than calling `pip` directly, because it guarantees that the package is installed into the **currently active interpreter**. This is particularly important when multiple Python installations or multiple virtual environments exist on the same machine.

### Install Pandoc

Install the **Pandoc** binary and ensure it is available in your system `PATH`.

This script does not rely only on the Python wrapper. It also probes the external `pandoc` executable, for example when querying version information and built-in templates. Therefore, installing `pypandoc` alone is **not sufficient**; the actual `pandoc` program must also be available.

### Install a LaTeX engine

Install at least one of the following and ensure it is available in `PATH`:

- `xelatex`
- `lualatex`
- `pdflatex`

The script automatically selects the first available engine in this order:

1. `xelatex`
2. `lualatex`
3. `pdflatex`

### Confirm the environment is correct

After activation and package installation, confirm that the project is using the correct Python and pip:

```bat
where python
where pip
python -V
python -m pip -V
```

On Linux or macOS:

```bash
which python
which pip
python -V
python -m pip -V
```

Also confirm that the external tools are visible:

```bash
pandoc --version
xelatex --version
```

If `xelatex` is not installed, the script will try `lualatex`, then `pdflatex`.

### First runtime test

Before attempting an executable build, run the script once from the activated virtual environment:

```bash
python script.py
```

or:

```bash
python script.py README.md
```

If the script starts correctly here, the Python-side dependencies are in place and you can proceed to executable packaging.

---

## Standalone executable build with PyInstaller

This project is a single-script application, so **PyInstaller** is the most direct way to produce a standalone executable.

### Important note on the tool name

The packaging tool is **PyInstaller**. If you wrote or meant “pycompiler”, use **PyInstaller** in the actual commands.

### Install PyInstaller into the same virtual environment

With the virtual environment activated:

```bash
python -m pip install pyinstaller
```

### Build a single-file Windows executable

For the GUI application form, use:

```bat
python -m PyInstaller --clean --onefile --windowed --name PandocMarkdownConverter script.py
```

This produces a single executable in:

```text
dist\PandocMarkdownConverter.exe
```

### Build a console executable instead

If you want a terminal window to remain visible for debugging, omit `--windowed`:

```bat
python -m PyInstaller --clean --onefile --name PandocMarkdownConverter script.py
```

### Recommended clean build sequence on Windows

```bat
cd C:\path\to\your\project
rmdir /s /q build
rmdir /s /q dist
del /q PandocMarkdownConverter.spec
py -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip setuptools wheel
python -m pip install pypandoc pypdf pyinstaller
python script.py --list-templates
python -m PyInstaller --clean --onefile --windowed --name PandocMarkdownConverter script.py
```

### Files and folders created by PyInstaller

After a build, you will typically see:

```text
build/
dist/
PandocMarkdownConverter.spec
```

The file you distribute is the executable inside `dist/`.

### Important packaging note for this project

PyInstaller bundles the Python interpreter and imported Python modules, but it does **not** automatically replace external system tools that your script expects at runtime.

This script still requires the following to be available on the target machine unless you redesign the program to bundle or locate them differently:

- `pandoc`
- at least one LaTeX engine such as `xelatex`, `lualatex`, or `pdflatex`

So the executable is “standalone” with respect to Python, but the full conversion workflow still depends on those external document-conversion tools.

### When to recreate the virtual environment

If the project folder, parent folder, or `.venv` folder has been moved or renamed, recreate the virtual environment instead of trying to reuse it. Virtual environments are not reliably portable after path changes.

---

## Quick start

### Start the GUI

```bash
python script.py
```

### Convert a file from the command line

```bash
python script.py README.md
```

### Specify explicit output names

```bash
python script.py README.md --docx output.docx --pdf output.pdf
```

### Force GUI mode

```bash
python script.py --gui
```

### List discoverable templates

```bash
python script.py --list-templates
```

---

## GUI usage

When the script is started with no command-line arguments, it opens a Tkinter-based graphical interface.

### GUI layout

The GUI contains three main tabs:

- **Convert**
- **Log**
- **Instructions**

### Convert tab

The **Convert** tab is divided into several areas.

#### Files section

This section lets you define:

- **Input Markdown** file
- **Output DOCX** file
- **Output PDF** file

The output names automatically track the selected input file until you manually replace them.

#### Pandoc templates section

This section contains:

- a read-only dropdown listing discovered templates
- a **Refresh templates** button
- a short summary of the selected template

The default selection logic prefers the explicit `default` template.

#### Actions section

The interface includes these buttons:

- **Convert now**
- **Open Log**
- **Open Instructions**
- **Quit**

#### Current conversion profile

The right-side summary panel shows the current PDF and math conversion profile, including:

- A4 paper
- no side panels
- one-column view
- fit-to-width opening
- LaTeX engine priority
- math header package set

### Log tab

The **Log** tab shows runtime messages such as:

- selected PDF engine
- conversion stage progress
- template fallback notes
- error messages

### Instructions tab

The **Instructions** tab contains a built-in overview of the script's purpose and usage.

---

## CLI usage

The CLI mode is used whenever you provide input arguments other than a plain GUI start.

### Basic form

```bash
python script.py <input_markdown>
```

### With explicit output paths

```bash
python script.py <input_markdown> --docx <output_docx> --pdf <output_pdf>
```

### With a template selection

```bash
python script.py <input_markdown> --template latex
```

The `--template` argument accepts either:

- a `.latex` file path
- a built-in Pandoc format name such as `latex`

### Listing templates

```bash
python script.py --list-templates
```

This prints the discovered template labels and exits.

---

## Command-line arguments

The script defines the following CLI interface.

### Positional argument

#### `input`

Input Markdown file.

Default:

```text
README.md
```

### Optional arguments

#### `--docx`

Explicit DOCX output path.

If omitted, the path is derived from the input filename.

#### `--pdf`

Explicit PDF output path.

If omitted, the path is derived from the input filename.

#### `--template`

PDF template selector.

Accepted forms:

- path to a `.latex` template file
- built-in format name such as `latex`

#### `--list-templates`

Lists discoverable templates and exits.

#### `--gui`

Forces GUI mode.

---

## Examples

### Example 1: Convert the default `README.md`

```bash
python script.py README.md
```

Expected outputs:

```text
README.docx
README.pdf
```

### Example 2: Open the GUI

```bash
python script.py
```

### Example 3: Convert with explicit paths

```bash
python script.py report.md --docx dist/report.docx --pdf dist/report.pdf
```

### Example 4: Use a built-in LaTeX template route

```bash
python script.py report.md --template latex
```

### Example 5: List available template labels

```bash
python script.py --list-templates
```

---

## Project structure

This repository can be as simple as a single-script project:

```text
.
├── script.py
└── README.md
```

The script dynamically creates temporary files during execution, including:

- a temporary preprocessed Markdown file
- a temporary LaTeX header file
- sometimes a temporary exported built-in template file

These temporary files are cleaned up automatically.

---

## Key implementation details

This section summarises the main internal functions and their roles.

### `preprocess_markdown(content)`

Normalises Markdown before conversion.

Responsibilities include:

- line ending normalisation
- math block cleanup
- display-math conversion
- targeted formula replacements

### `convert_bracket_display_math(content)`

Converts:

```latex
\[ ... \]
```

into:

```latex
$$ ... $$
```

### `convert_latex_display_environments(content)`

Transforms supported display math environments into Pandoc-friendly display blocks.

### `choose_pdf_engine()`

Selects the first available engine in this order:

1. `xelatex`
2. `lualatex`
3. `pdflatex`

### `discover_builtin_templates()`

Queries Pandoc built-in templates and keeps only PDF-compatible ones.

### `discover_custom_templates()`

Scans standard template folders for custom `.latex` templates.

### `discover_pandoc_templates()`

Builds the final list of selectable templates, including:

- explicit `default`
- built-in PDF-compatible templates
- custom PDF-compatible templates

### `build_pdf_template_args(template_info)`

Builds the Pandoc `--template` arguments for PDF generation.

### `convert_markdown_file(...)`

Main conversion routine.

Responsibilities include:

- resolving input/output paths
- reading and preprocessing Markdown
- generating DOCX
- generating PDF
- retrying with default template if needed
- applying PDF viewer settings

### `apply_pdf_view_settings(pdf_file)`

Rewrites the generated PDF with viewer preferences.

### `PandocConverterGUI`

Tkinter application class that builds and controls the GUI.

---

## Error handling and fallbacks

The script includes several defensive behaviours.

### Missing Pandoc

If Pandoc is not found, the script raises a clear runtime error.

### Missing LaTeX engine

If none of `xelatex`, `lualatex`, or `pdflatex` is found, PDF generation cannot proceed and the script reports the issue.

### Invalid input file

If the input Markdown file does not exist, the script stops with a clear error.

### Incompatible template

If a selected template is not PDF-compatible, the script falls back to the default Pandoc PDF template route.

### Missing `.cls` or `.sty` files

If a selected LaTeX template fails because required LaTeX class or style files are missing, the script retries PDF generation with the default template.

### Tkinter unavailable

If GUI mode is requested but Tkinter is not available in the Python installation, the script reports the issue and exits.

---

## Troubleshooting

### Pandoc not found

Symptom:

```text
Pandoc was not found
```

Action:

- install Pandoc
- confirm `pandoc --version` works in your terminal
- ensure Pandoc is in `PATH`

### PDF engine not found

Symptom:

```text
No supported LaTeX PDF engine found
```

Action:

- install `xelatex`, `lualatex`, or `pdflatex`
- confirm at least one of them is accessible from `PATH`

### GUI does not open

Symptom:

```text
Tkinter is not available
```

Action:

- install a Python distribution that includes Tkinter
- or use the script in CLI mode

### Selected template fails during PDF build

Possible cause:

- the template depends on missing `.cls` or `.sty` files

Action:

- use the default template
- install the missing LaTeX package(s)
- simplify the chosen custom template

### `ModuleNotFoundError: No module named 'pypandoc'`

Symptom:

```text
Traceback (most recent call last):
  File "script.py", line 151, in <module>
    import pypandoc
ModuleNotFoundError: No module named 'pypandoc'
```

Meaning:

- the active Python interpreter does not have `pypandoc` installed
- or the terminal is using the wrong Python interpreter
- or the virtual environment was not activated correctly

Action:

1. Activate the correct virtual environment.
2. Install the package through that interpreter:

```bash
python -m pip install pypandoc pypdf
```

3. Confirm installation:

```bash
python -m pip show pypandoc
python -c "import pypandoc; print('pypandoc OK')"
```

### `pip` launcher points to the wrong virtual environment

Symptom on Windows:

```text
Fatal error in launcher: Unable to create process ...
```

Typical cause:

- the `.venv` folder was copied, moved, or the parent project directory was renamed
- `pip.exe` still points to an older Python path

Action:

Delete and recreate the virtual environment:

```bat
rmdir /s /q .venv
py -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip setuptools wheel
python -m pip install pypandoc pypdf pyinstaller
```

### PyInstaller build succeeds but conversion fails on another machine

Possible cause:

- Python was bundled correctly
- but `pandoc` and/or the LaTeX engine are missing on the target machine

Action:

- install **Pandoc**
- install **MiKTeX**, **TeX Live**, or another LaTeX distribution providing `xelatex`, `lualatex`, or `pdflatex`
- confirm the tools are in `PATH`

### Formulas render poorly in DOCX

Possible causes:

- the original Markdown contains unsupported raw LaTeX structures
- the input uses uncommon macros not understood by Pandoc or Word

Action:

- prefer standard Pandoc math syntax
- prefer `$$ ... $$` display math where possible
- reduce highly custom raw LaTeX when DOCX fidelity is critical

---

## Limitations

This script is practical and robust for many use cases, but some boundaries remain.

### DOCX math is not full LaTeX

Pandoc converts compatible math to Office Math, but not every LaTeX construct has a perfect Word equivalent.

### Template compatibility is intentionally restricted

The current implementation only treats LaTeX-oriented templates as PDF-compatible. This is deliberate, because the PDF route in this tool is LaTeX-based.

### PDF viewer preferences are advisory

Not all PDF viewers honour embedded opening preferences in the same way.

### Very custom LaTeX documents may still require manual tuning

If your Markdown relies on:

- external class files
- complex macro packages
- custom environments
- advanced TeX programming

then you may still need to adapt the source or the template.

---

## Recommended use cases

This script is particularly suitable for:

- scientific Markdown notes
- engineering reports
- technical memos
- formula-heavy documentation
- project documentation that must be exported to both Word and PDF
- users who want a GUI but still need a CLI option

It is especially useful when you want:

- a simple single-file tool
- predictable output naming
- A4 PDF output
- better control over PDF opening behaviour
- a safer path for LaTeX-style equations in both DOCX and PDF

---

## Summary

`Pandoc Markdown Converter` is a practical bridge between Markdown authoring and dual-format publication.

Its main strengths are:

- direct DOCX + PDF generation from one Markdown source
- math-aware preprocessing
- LaTeX-backed PDF rendering
- template discovery
- explicit default template selection
- post-processed PDF viewing preferences
- both GUI and CLI workflows in one script

For users working with technical Markdown, this provides a more structured and dependable conversion workflow than a minimal Pandoc command wrapper.