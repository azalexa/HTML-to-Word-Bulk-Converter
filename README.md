# HTML to Word Bulk Converter

PowerShell automation tool for bulk HTML to Word conversion.

## Features

* HTML and HTM support
* DOC and DOCX conversion
* Image embedding
* Missing image detection
* CSV reporting
* Progress tracking
* Automatic path detection
* Folder structure preservation
* Duplicate document protection
* Error handling

## Requirements

* Windows
* PowerShell 5.1
* Microsoft Word

## Usage

Place ConvertHTML.ps1 in the folder containing HTML files.

Run ConvertHTML.ps1 using PowerShell.

Converted documents and reports are saved in the Converted folder.

## Limitations

* Microsoft Word must be installed.
* Advanced HTML formatting may not be preserved.
* Internal article links are not automatically remapped.
* External images may require authentication.
