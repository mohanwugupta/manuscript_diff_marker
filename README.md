# color_docx_changes

A Python command-line tool that highlights text changes between two `.docx` (or legacy `.doc`) files by coloring the changed words directly in the revised document — without using Word's Track Changes feature.

Designed for researchers who need a clean, portable marked-up manuscript that displays correctly in any PDF viewer or submission system.

---

## What it does

- Compares an **original** and a **revised** `.docx` at the paragraph and word level.
- Outputs a copy of the revised document with **changed/inserted text colored** (default: red).
- Leaves deleted text (present only in the original) unmarked — the output is always a clean version of the revised document.
- Skips **formatting-only changes** (same words, different bold/font/size) — only actual text edits are highlighted.
- Preserves all run-level formatting (bold, italic, superscript, font size, etc.) and never displaces inline elements such as citation superscripts, hyperlinks, bookmarks, or field codes.

---

## Requirements

- Python 3.9 or later (tested on 3.13)
- [`python-docx`](https://python-docx.readthedocs.io/)
- *(Legacy `.doc` files only)* Microsoft Word installed + [`pywin32`](https://pypi.org/project/pywin32/) (Windows only)

---

## Installation

```bash
# 1. Clone the repository
git clone https://github.com/YOUR_USERNAME/color_docx_changes.git
cd color_docx_changes

# 2. Create and activate a virtual environment (recommended)
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS / Linux:
source .venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt
```

---

## Usage

```bash
python color_docx_changes.py \
  --original "path/to/original.docx" \
  --revised  "path/to/revised.docx" \
  --out      "path/to/revised_marked.docx" \
  --rgb      "CC0000"
```

| Argument | Required | Description |
|---|---|---|
| `--original` | Yes | Path to the original (baseline) `.docx` |
| `--revised` | Yes | Path to the revised `.docx` |
| `--out` | Yes | Output path for the marked-up document |
| `--rgb` | No | Hex color for changed text (default: `CC0000` = red) |

Both `.docx` and legacy `.doc` files are accepted. `.doc` files are auto-converted via Word COM (Windows + `pywin32` required).

### Example

```bash
python color_docx_changes.py \
  --original "manuscript_v1.docx" \
  --revised  "manuscript_v2.docx" \
  --out      "manuscript_v2_marked.docx" \
  --rgb      "CC0000"
```

---

## How it works

1. Both documents are parsed with `python-docx`.
2. Non-empty paragraphs are extracted and normalized, then aligned at the paragraph level using Python's `difflib.SequenceMatcher`.
3. For paired changed paragraphs, a second word-level diff identifies exactly which tokens changed.
4. Changed spans are colored by modifying `<w:r>` XML elements **in-place** — using `deepcopy` + `lxml`'s `addprevious()` when a run must be split. This ensures that non-run XML siblings (citation superscripts, hyperlinks, bookmarks, footnote references) are never reordered.

---

## Limitations

- Tables and text boxes are not currently diffed (only main-body paragraphs).
- Structural changes (many paragraphs reordered at once) may produce whole-paragraph coloring instead of word-level coloring.
- `.doc` auto-conversion requires Microsoft Word on Windows.

---

## License

MIT — see [LICENSE](LICENSE).

---

## Citation

If this tool is useful for your research, a mention in your acknowledgements is appreciated.
