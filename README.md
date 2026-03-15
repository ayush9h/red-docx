# reddocx


![PyPI version](https://img.shields.io/pypi/v/reddocx)
![Python versions](https://img.shields.io/pypi/pyversions/reddocx)
![License](https://img.shields.io/pypi/l/reddocx)
![Status](https://img.shields.io/badge/status-active-success)

Lightweight track-changes engine for Microsoft Word (.docx) files using direct XML manipulation.

`reddocx` provides a minimal and fast way to programmatically add tracked revisions
(insertions and deletions) to Word documents without requiring Microsoft Word or
heavy document processing libraries.


## Features

- Word-style tracked changes (insert / delete revisions)
- Paragraph-level word replacement tracking
- Pure XML processing using lxml
- No Microsoft Word dependency
- Lightweight and fast
- Supports file path, bytes, or memory streams


## Installation

 - Default Installation
```bash
pip install reddocx
 ```
 - UV installation
 ```bash
 uv pip install reddocx