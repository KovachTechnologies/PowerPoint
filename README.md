# PowerPoint Utility Toolkit

A lightweight, command-line Python script that provides multiple utilities for working with Microsoft PowerPoint (.pptx) files using the `python-pptx` library.

# Features

- Create a simple hardcoded example table slide (`create_simple_table`)
- Generate PowerPoint files containing tables extracted from Markdown files in a directory (`create_tables`)
- Convert structured text files (with titles + nested bullets) into nicely formatted PowerPoint presentations (`create_from_text`)
- Extract slide titles and bullet-point content from an existing .pptx file back into a structured text format (`extract_to_text`)

# Requirements

```text
python-pptx
```

Install with:

```bash
pip install python-pptx
```

or

```bash
pip install -r requirements.txt
```

# Installation

Just download or copy the single file:

```bash
# Recommended: place it somewhere in your PATH or use it directly
wget https://raw.githubusercontent.com/yourusername/powerpoint-utils/main/powerpoint.py
# or simply copy-paste into powerpoint.py
```

No further installation steps are needed beyond the dependency above.

# Usage

```bash
python powerpoint.py <mode> [options]
```

# Available Modes

## 1. Create a simple hardcoded table presentation

```bash
python powerpoint.py create_simple_table --output example-table.pptx
```

Creates a single-slide PPTX with a 4×3 demo table (Name / Age / City).

## 2. Create table slides from Markdown files

Processes every `.md` file in a directory and creates one `_tables.pptx` file per markdown file containing all detected tables.

```bash
python powerpoint.py create_tables --dir "Day 1"
# or
python powerpoint.py create_tables --dir "/path/to/markdown/folder"
```

Expected markdown format (example):

```markdown
#### Slide 3 – Important Numbers
| Metric       | Value    | Unit     |
|--------------|----------|----------|
| Throughput   | 1200     | req/s    |
| Latency p99  | 45       | ms       |
```

Produces `Day 1/0103_tables.pptx` (or similar)

## 3. Create presentation from structured text file

Input file format (`slides.txt` example):

```text
Slide 1
====================
Title: Introduction

Body Bullets:
- Welcome to the project
  - Goal 1
  - Goal 2

Slide 2
====================
Title: Architecture

Body Bullets:
- Component A
  - Sub-component A1
- Component B
```

Command:

```bash
python powerpoint.py create_from_text --input slides.txt --output presentation.pptx
# or just:
python powerpoint.py create_from_text --input slides.txt
# → creates slides_recreated.pptx
```

## 4. Extract bullets and titles from an existing PPTX

```bash
python powerpoint.py extract_to_text \
  --input presentation.pptx \
  --output extracted-slides.txt
```

Produces a text file very similar to the format expected by `create_from_text`.

## Example Workflow

1. Write lecture notes or meeting agendas in Markdown with tables
2. Run `create_tables` → get one PPTX per file with clean tables
3. Create bullet-point slides in plain text → convert to PPTX with `create_from_text`
4. After presenting → extract content back to text for editing/version control

## Project Structure Suggestion

```
powerpoint-utils/
├── powerpoint.py           # the main script (single file)
├── README.md
├── requirements.txt
├── example/
│   ├── slides.txt
│   ├── presentation.pptx
│   └── extracted-slides.txt
└── Day 1/                  # example markdown folder
    ├── 0101.md
    └── 0101_tables.pptx
```

# License

MIT License (feel free to use, modify, distribute)

# Contributing

Bug reports, table parsing improvements, better bullet indentation handling, support for speaker notes, images, etc. are welcome!

Happy presenting! 🚀
