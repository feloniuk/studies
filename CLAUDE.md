# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

This is a university study materials repository with three courses:
- **1_Диверсифікація інвестицій** — Investment Diversification (lectures in PDF, assignments in .docx)
- **2_Тайм менеджмент** — Time Management (lectures in .pptx/.docx, assignments in .docx)
- **3_Планування інноваційної діяльності** — Innovation Planning (lectures in PDF, assignments in .docx)

Each course is organized by modules (Модуль 1, 2, ...). Lecture files contain theory; practical assignment files (Практичне завдання) contain tasks to be solved.

## Working with files

### Reading .docx files
```python
py -3 -c "
from docx import Document
doc = Document('path/to/file.docx')
for p in doc.paragraphs:
    print(p.text)
"
```

### Reading PDF files
```python
py -3 -c "
import fitz
doc = fitz.open('path/to/file.pdf')
for page in doc:
    print(page.get_text())
"
```

### Writing answers to .docx files
```python
py -3 -c "
from docx import Document
doc = Document('path/to/task.docx')
# Add paragraphs with answers
doc.add_paragraph('Answer text here')
doc.save('path/to/task.docx')
"
```

Use `python-docx` to preserve existing content and append/modify answers.

### Converting with pandoc
```bash
pandoc "file.docx" -t plain  # quick text extraction
pandoc "file.pdf" -t plain   # PDF to text
```

## Installed tools

- **Python**: `py -3` (Python 3.14, via `py` launcher — `python` and `python3` don't work on this system)
- **python-docx**: read/write .docx files
- **PyMuPDF (fitz)**: read PDF files
- **PyPDF2**: alternative PDF reader
- **pandoc 3.9**: document conversion

## Workflow for solving assignments

1. Read the assignment .docx file to understand the task
2. Read the relevant lecture PDF/docx files for theory
3. Compose the answer based on lecture content
4. Write the answer back into the assignment .docx file (or create a new one)

## Language

All materials are in Ukrainian. Answers should be written in Ukrainian unless specified otherwise.
