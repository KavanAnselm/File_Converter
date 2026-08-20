# File Converter

[![Python](https://img.shields.io/badge/Python-3.x-blue)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/Pandas-Data%20Processing-150458)](https://pandas.pydata.org/)
[![Pillow](https://img.shields.io/badge/Pillow-Image%20Processing-yellow)](https://python-pillow.org/)
[![MoviePy](https://img.shields.io/badge/MoviePy-Video%20Processing-green)](https://zulko.github.io/moviepy/)
[![PyDub](https://img.shields.io/badge/PyDub-Audio%20Processing-orange)](https://github.com/jiaaro/pydub)

**File Converter** is a command-line Python utility for converting and extracting content across several common file formats. It supports audio, image, video, spreadsheet, PDF, and DOCX-to-text operations through a simple menu-driven interface.

The program validates the input file, presents format choices where applicable, performs the requested conversion, and saves the converted file beside the original file using the new extension. fileciteturn2file2L78-L105

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Supported Conversions](#supported-conversions)
- [Tech Stack](#tech-stack)
- [Project Structure](#project-structure)
- [Architecture Diagram](#architecture-diagram)
- [Project Flow Diagram](#project-flow-diagram)
- [Key Components](#key-components)
- [Setup & Installation](#setup--installation)
- [Running the Converter](#running-the-converter)
- [Usage](#usage)
- [Output Naming](#output-naming)
- [Troubleshooting](#troubleshooting)
- [Potential Improvements](#potential-improvements)

---

## Overview

The application provides a single terminal menu for multiple file-conversion operations.

Available options are:

```text
1. Convert Audio
2. Convert Image
3. Convert Video
4. Convert CSV to XLSX
5. Convert XLSX to CSV
6. Extract Text from PDF
7. Convert DOCX to TXT
8. Exit
```

These options are implemented directly in the program's main menu. fileciteturn2file2L85-L100

---

## Features

### Audio Conversion

Supports input formats:

```text
MP3
WAV
OGG
FLAC
AAC
```

and allows the user to select an output format from the same set.

Audio conversion is implemented with PyDub's `AudioSegment`. fileciteturn2file2L9-L16

### Image Conversion

Supports:

```text
JPEG
JPG
PNG
BMP
GIF
TIFF
```

Pillow is used to open and save image files. fileciteturn2file2L17-L25

### Video Conversion

Supports input formats:

```text
MP4
AVI
MOV
MKV
WMV
```

MoviePy is used for video loading and export. The current implementation writes the output using the `libx264` codec. fileciteturn2file2L27-L34

### CSV to XLSX

Pandas reads the CSV and writes an Excel workbook.

### XLSX to CSV

Pandas reads the Excel workbook and writes a CSV file. fileciteturn2file2L35-L52

### PDF to TXT

PyPDF2 extracts text from every page and saves it as a `.txt` file. fileciteturn2file2L54-L65

### DOCX to TXT

Python-docx reads the document paragraphs and writes them into a text file. fileciteturn2file2L67-L76

---

## Supported Conversions

| Input | Output |
|---|---|
| MP3 / WAV / OGG / FLAC / AAC | MP3 / WAV / OGG / FLAC / AAC |
| JPEG / JPG / PNG / BMP / GIF / TIFF | JPEG / JPG / PNG / BMP / GIF / TIFF |
| MP4 / AVI / MOV / MKV / WMV | MP4 / AVI / MOV / MKV / WMV |
| CSV | XLSX |
| XLSX | CSV |
| PDF | TXT |
| DOCX | TXT |

---

## Tech Stack

| Layer | Technology |
|---|---|
| Programming Language | Python |
| Audio Processing | PyDub |
| Image Processing | Pillow |
| Video Processing | MoviePy |
| Spreadsheet Processing | Pandas |
| PDF Processing | PyPDF2 |
| DOCX Processing | python-docx |
| Interface | Command Line |

---

## Project Structure

```text
File_Converter/
│
├── File_Converter.py       # Main conversion utility
└── README.md               # Project documentation
```

---

## Architecture Diagram

```text
+---------------------------+
|       User Terminal       |
+-------------+-------------+
              |
              v
+---------------------------+
|      Main Menu            |
|   Select Conversion Type  |
+-------------+-------------+
              |
              v
+---------------------------+
|       File Path           |
|      Validation           |
+-------------+-------------+
              |
       +------+------+
       |             |
       v             v
 Supported       Unsupported
  Format           Format
       |             |
       v             v
+-------------+   Error Message
| Conversion  |
| Function    |
+------+------+ 
       |
       +------------------------------+
       |       |       |       |      |
       v       v       v       v      v
     Audio   Image   Video   Data   Documents
       |       |       |       |      |
       +-------+-------+-------+------+
                       |
                       v
              +----------------+
              | Output File    |
              +----------------+
```

---

## Project Flow Diagram

```text
Start
  │
  ▼
Display Conversion Menu
  │
  ▼
Select Option
  │
  ├── Exit ─────────────► End
  │
  ▼
Enter File Path
  │
  ▼
Check File Exists
  │
  ├── No ───────────────► Error → Menu
  │
  ▼
Validate File Extension
  │
  ├── Invalid ──────────► Error → Menu
  │
  ▼
Choose Output Format
  │
  ▼
Run Conversion
  │
  ▼
Save Output Beside Source
  │
  ▼
Display Result
  │
  ▼
Return to Menu
```

---

## Key Components

### `convert_audio()`

Loads audio through PyDub and exports it using the selected format.

### `convert_image()`

Uses Pillow's `Image.open()` and `Image.save()` methods to convert image formats.

### `convert_video()`

Loads a video with MoviePy and exports it with the `libx264` codec.

### `convert_csv_to_xlsx()`

Reads a CSV with Pandas and exports an `.xlsx` file.

### `convert_xlsx_to_csv()`

Reads an Excel workbook with Pandas and exports a `.csv` file.

### `extract_text_from_pdf()`

Iterates over PDF pages using `PdfReader`, extracts their text, and saves the result as `.txt`.

### `convert_docx_to_txt()`

Reads all DOCX paragraphs and joins them with newline characters.

### `get_conversion_choice()`

Displays numbered output-format choices and returns the user's selection.

### `main()`

Controls the menu-driven application loop, validates file paths and extensions, and dispatches each operation to the appropriate conversion function. fileciteturn2file2L85-L134

---

## Setup & Installation

### 1. Clone the repository

```bash
git clone https://github.com/KavanAnselm/File_Converter.git
cd File_Converter
```

### 2. Create a virtual environment

```bash
python -m venv venv
```

Activate it:

#### Windows

```bash
venv\Scripts\activate
```

#### Linux/macOS

```bash
source venv/bin/activate
```

### 3. Install dependencies

```bash
pip install pandas pydub Pillow moviepy PyPDF2 python-docx
```

Some audio/video conversions may additionally require system-level codecs or FFmpeg depending on the operating system and installed backend.

---

## Running the Converter

```bash
python File_Converter.py
```

The application displays:

```text
File Converter
1. Convert Audio
2. Convert Image
3. Convert Video
4. Convert CSV to XLSX
5. Convert XLSX to CSV
6. Extract Text from PDF
7. Convert DOCX to TXT
8. Exit
```

Choose an operation and provide the path to the input file. fileciteturn2file2L85-L105

---

## Usage

### Convert Audio

Select:

```text
1
```

Then choose an output format:

```text
1. mp3
2. wav
3. ogg
4. flac
5. aac
```

### Convert Image

Select:

```text
2
```

Then choose the desired image format.

### Convert Video

Select:

```text
3
```

Then choose the output video format.

### Convert CSV to XLSX

Select:

```text
4
```

Provide the path to a `.csv` file.

### Convert XLSX to CSV

Select:

```text
5
```

Provide the path to an `.xlsx` file.

### Extract Text from PDF

Select:

```text
6
```

Provide a `.pdf` file.

### Convert DOCX to TXT

Select:

```text
7
```

Provide a `.docx` file.

---

## Output Naming

The converter normally creates the output file in the same directory as the input file.

For example:

```text
report.csv
```

becomes:

```text
report.xlsx
```

and:

```text
document.pdf
```

becomes:

```text
document.txt
```

The output path is generated by replacing the original extension with the requested output extension. fileciteturn2file2L9-L14

---

## Troubleshooting

### File does not exist

The application checks the supplied path with `os.path.isfile()` before conversion. fileciteturn2file2L103-L106

### Unsupported format

Make sure the selected conversion matches the input extension supported by that menu option.

### Invalid output selection

The application handles invalid numeric choices through `ValueError` and `IndexError` handling.

### Video conversion fails

MoviePy video conversion can depend on FFmpeg and the codecs available on the system.

### Audio conversion fails

PyDub may require an appropriate FFmpeg/audio backend depending on the source and target formats.

### PDF extraction is incomplete

PyPDF2 extracts embedded PDF text. Scanned/image-only PDFs may require OCR.

---

## Potential Improvements

- Add a graphical interface.
- Add drag-and-drop file support.
- Add batch conversion.
- Add progress indicators for large media files.
- Add automatic output-directory selection.
- Add better validation for output-format choices.
- Add a `requirements.txt`.
- Add logging to a file.
- Add overwrite confirmation when an output file already exists.
- Add OCR support for scanned PDFs.
- Add more document conversions such as DOCX ↔ PDF.
- Improve video codec selection based on the requested output format.
- Add unit tests for each conversion function.

---

## License

© 2026. All rights reserved.

This project is closed source and may not be redistributed or modified without explicit permission from the maintainers.