# Prepare Emails for LLM

Convert email archives (EML files) into LLM-friendly formats, optimized for tools like NotebookLM.

## Overview

This utility processes directories of EML email files and converts them into a structured format that's perfect for ingestion into Large Language Models (LLMs). It:

- Parses EML files and extracts all content
- Recursively extracts and parses nested EML attachments
- Converts HTML emails to clean Markdown
- Consolidates all emails into a single, searchable markdown file
- Extracts and organizes attachments
- Generates cross-references between emails and attachments
- Creates a manifest for easy navigation

## Features

- **Comprehensive Email Parsing**
  Extracts headers, body (text/HTML), attachments, and metadata

- **Recursive Nested Email Handling**
  Automatically detects and parses .eml attachments, inline with the parent email

- **LLM-Optimized Output**
  Generates clean Markdown with proper formatting and structure

- **Attachment Management**
  Saves attachments with unique IDs and creates references in the markdown

- **NotebookLM Ready**
  Output format designed to work within NotebookLM's 50-file limit

## Installation

1. **Create virtual environment:**
   ```bash
   mkvirtualenv prepare-emails-for-llm -p python3.11
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Basic Usage

```bash
python prepare-for-llm.py --input <email_directory> --output <output_directory>
```

### Example

```bash
python prepare-for-llm.py --input ../outlook-downloader/emails_comprehensive --output prepared-for-llm
```

### Command-line Options

- `--input`, `-i` (required) - Input directory containing EML files
- `--output`, `-o` (default: `prepared-for-llm`) - Output directory
- `--no-nested` - Don't extract nested .eml attachments
- `--verbose`, `-v` - Enable verbose logging

## Output Structure

```
prepared-for-llm/
├── README.md          # Summary and instructions
├── emails.md          # All emails in chronological order (MAIN FILE)
├── manifest.json      # Machine-readable metadata
└── attachments/       # All non-EML attachments
    ├── att-0001-001_document.pdf
    ├── att-0002-001_image.png
    └── ...
```

### emails.md Format

The main markdown file contains all emails in chronological order:

```markdown
# Email Archive

**Total Emails:** 303
**Generated:** 2024-11-27 22:30:00

---

## Email 1: Meeting Tomorrow

**Metadata:**
- **From:** John Doe <john@example.com>
- **To:** jane@example.com
- **Date:** 2024-01-15 10:30:00
- **Attachments:** 2
  - `agenda.pdf` (application/pdf, 45.2 KB) [→ att-0001-001]
  - `slides.pptx` (application/vnd..., 1.2 MB) [→ att-0001-002]

**Message:**

Hi Jane,

Looking forward to our meeting tomorrow...

---

## Email 2: Follow-up

...
```

### Nested Email Handling

When an email contains .eml attachments, they are parsed and displayed inline:

```markdown
## Email 5: Forwarded Message

**Metadata:**
- **From:** alice@example.com
- **Attachments:** 1
  - `original-email.eml` (email message, parsed below)

**Message:**

See the email below...

### Attached Emails

#### Attached Email 1: Original Subject

- **From:** Bob Smith <bob@example.com>
- **Date:** 2024-01-10 14:20:00

Original message content here...
```

## Using with NotebookLM

NotebookLM has a 50-file upload limit. Here's how to use the output:

1. **Upload the main file:**
   - Upload `emails.md` (contains all email text)

2. **Select key attachments:**
   - Review the `attachments/` directory
   - Upload up to 49 most important attachments
   - NotebookLM can reference these in responses

3. **Query your emails:**
   - Ask questions about your email archive
   - NotebookLM will search across all content

### Example Queries

- "What did John say about the project timeline?"
- "Find all emails related to the contract"
- "Summarize the conversation about budget"
- "What attachments were sent in emails about the proposal?"

## Technical Details

### Email Parsing

The tool uses Python's `email` library with the following features:

- **Header Decoding:** Properly handles encoded headers (RFC 2047)
- **Charset Detection:** Automatically detects and converts character encodings
- **HTML Conversion:** Converts HTML emails to clean Markdown using `html2text`
- **Attachment Extraction:** Handles base64 and other encodings

### Recursive Processing

When `.eml` files are found as attachments:

1. The attachment is detected by extension or MIME type
2. The EML is parsed as a complete email
3. Its content is included inline in the markdown
4. Its attachments are also extracted and saved

This process is recursive, so emails-within-emails-within-emails are fully supported.

### Attachment ID System

Each attachment gets a unique ID in the format:

- `att-<email_number>-<attachment_number>` for direct attachments
- `att-<email_number>-<nested_number>-<attachment_number>` for nested attachments

Example: `att-0042-002` = 2nd attachment of email #42

## Testing the Parser

You can test the EML parser independently:

```bash
python eml_parser.py path/to/email.eml
```

This will display:
- Subject, From, To, Date
- List of attachments
- Nested emails (if any)
- First 500 characters of the body

## Troubleshooting

### "No emails found to process"

- Check that the input directory contains `.eml` files
- Try using `--verbose` to see what files are being scanned

### "Error parsing email"

- Some EML files may be corrupted or use unusual encoding
- Check the logs for specific error messages
- The tool will skip problematic emails and continue processing

### Large Output Files

- For very large email archives (1000+ emails), consider splitting by date
- The markdown file may be several MB - most LLMs can handle this

## Limitations

- **Binary Attachments:** Images, PDFs, etc. are saved as separate files
- **Formatting:** Some complex HTML formatting may not convert perfectly to Markdown
- **Embedded Images:** Inline images are treated as attachments
- **File Size:** Very large attachments (100MB+) may cause memory issues

## License

This is a utility script for personal/organizational use.

## Contributing

Feel free to submit issues or pull requests for improvements.
