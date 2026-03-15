# reddocx


![PyPI version](https://img.shields.io/pypi/v/reddocx)
![Python versions](https://img.shields.io/pypi/pyversions/reddocx)
![License](https://img.shields.io/pypi/l/reddocx)
![Status](https://img.shields.io/badge/status-active-success)

Lightweight track-changes engine for Microsoft Word (.docx) files using direct XML manipulation.

`reddocx` provides a minimal and fast way to programmatically add tracked revisions
(insertions and deletions) to Word documents without requiring Microsoft Word or
heavy document processing libraries.

## Installation

 - Default Installation
```bash
pip install reddocx
 ```
 - UV installation
 ```bash
 uv pip install reddocx
```


## Usage
```bash
from reddocx.core.document import DocxDocument
doc = DocxDocument('sample.docx')
updated_report = doc.track_replace_words({'original_word':'replaced_word'})
updated_doc = doc.save()
```

## Demo

### Before → Original document

<p align="center">
  <img src="https://drive.google.com/uc?export=view&id=1RaZM0_r6Hmt16yZ078J95fMOal33tatv" width="900"/>
</p>

### After → Track changes applied by `reddocx`

<p align="center">
  <img src="https://drive.google.com/uc?export=view&id=1I99g2VKfX9OmvBmHFdlTpSj2nZ9w8Z9Y" width="900"/>
</p>



## Features

- Word-style tracked changes (insert / delete revisions)
- Paragraph-level word replacement tracking
- Pure XML processing using lxml
- No Microsoft Word dependency
- Lightweight and fast
- Supports file path, bytes, or memory streams

