# ppt-translator

A small Python utility to translate Japanese text in PowerPoint presentations to English using `python-pptx` and `deep-translator`.

## Features

- Translate Japanese text in `.pptx` files
- Skip image objects and only translate text content
- Avoid duplicate translations if there are both English and Japanese sentences on the slide

## Prerequisites

- Python 3.11 or later
- [Pixi](https://github.com/microsoft/pixi) for dependency and runtime management

## Installation

From the `ppt-translator/ppt-translator` directory:

```powershell
pixi install
```

## Usage

Put your ppt file in the ppt folder and run the translator with a source PowerPoint filename and optional output path:

```powershell
pixi run python .\src\ppt_translator\__init__.py .\ppt\<INPUT FILENAME>.pptx --output .\ppt\<OUTPUT FILENAME>.pptx
```

If `--output` is omitted, the translated file is saved alongside the source with `_translated` appended to the filename.

## Project Structure

- `pyproject.toml` — project configuration and dependencies
- `src/ppt_translator/__init__.py` — translation implementation
- `ppt/` — input/output PowerPoint files (ignored by Git)
