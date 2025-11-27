#!/usr/bin/env python3
"""
Prepare Emails for LLM (Enhanced) - Convert email archives with text extraction

This enhanced version extracts text from PDFs and DOCX files and includes them
inline in the markdown, reducing the number of separate files needed.
"""

import argparse
import json
import logging
from pathlib import Path
from typing import List, Dict, Tuple
from datetime import datetime
from eml_parser import EMLParser, ParsedEmail, EmailAttachment
from attachment_extractor import AttachmentExtractor
import re
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from PIL import Image
import io

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class EnhancedEmailProcessor:
    """Process emails with text extraction and bundling"""

    def __init__(self, input_dir: Path, output_dir: Path, extract_nested: bool = True):
        """
        Initialize the enhanced email processor

        Args:
            input_dir: Directory containing EML files
            output_dir: Directory for output files
            extract_nested: Whether to recursively extract nested EML attachments
        """
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.parser = EMLParser(extract_nested_emails=extract_nested)
        self.extractor = AttachmentExtractor()
        self.emails: List[ParsedEmail] = []
        self.attachment_counter = 0
        self.attachment_registry: Dict[str, str] = {}
        self.images_for_bundle: List[Tuple[bytes, str]] = []  # (content, filename)
        self.binary_files_for_bundle: List[Tuple[bytes, str, str]] = []  # (content, filename, type)

    def process_all_emails(self):
        """Process all EML files in the input directory"""
        logger.info(f"Scanning for EML files in {self.input_dir}")

        eml_files = list(self.input_dir.glob('**/*.eml'))
        logger.info(f"Found {len(eml_files)} EML files")

        for eml_file in eml_files:
            try:
                logger.debug(f"Processing: {eml_file.name}")
                parsed = self.parser.parse_file(eml_file)
                self.emails.append(parsed)
            except Exception as e:
                logger.error(f"Error processing {eml_file}: {e}")

        # Sort emails by date
        self.emails.sort(key=lambda e: e.date_obj if e.date_obj else datetime.min)

        logger.info(f"Successfully parsed {len(self.emails)} emails")

    def generate_markdown(self) -> str:
        """Generate markdown with extracted attachment text inline"""
        logger.info("Generating markdown with text extraction...")

        lines = []

        # Header
        lines.append("# Email Archive with Extracted Attachments")
        lines.append("")
        lines.append(f"**Total Emails:** {len(self.emails)}")
        lines.append(f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append("")
        lines.append("_This archive includes extracted text from PDF and DOCX attachments inline._")
        lines.append("")
        lines.append("---")
        lines.append("")

        # Process each email
        for idx, email in enumerate(self.emails, 1):
            lines.extend(self._format_email_with_extraction(email, idx))
            lines.append("")
            lines.append("---")
            lines.append("")

        return '\n'.join(lines)

    def _format_email_with_extraction(self, email: ParsedEmail, email_number: int) -> List[str]:
        """Format email with extracted attachment text"""
        lines = []

        # Header
        subject = email.subject or "(No Subject)"
        lines.append(f"## Email {email_number}: {subject}")
        lines.append("")

        # Metadata
        lines.append("**Metadata:**")
        lines.append(f"- **From:** {email.from_name} <{email.from_addr}>")
        if email.to_addrs:
            lines.append(f"- **To:** {', '.join(email.to_addrs)}")
        if email.cc_addrs:
            lines.append(f"- **Cc:** {', '.join(email.cc_addrs)}")
        lines.append(f"- **Date:** {email.date or 'Unknown'}")
        if email.message_id:
            lines.append(f"- **Message ID:** {email.message_id}")

        # Count attachments by type
        text_extracted = 0
        images = 0
        other = 0

        for att in email.attachments:
            if not att.is_eml:
                if self.extractor.can_extract_text(att.content_type, att.filename):
                    text_extracted += 1
                elif self.extractor.is_image(att.content_type, att.filename):
                    images += 1
                else:
                    other += 1

        if email.attachments:
            lines.append(f"- **Attachments:** {len(email.attachments)} ({text_extracted} text extracted, {images} images, {other} other)")

        lines.append("")

        # Body
        lines.append("**Message:**")
        lines.append("")
        if email.body_markdown:
            lines.append(email.body_markdown)
        else:
            lines.append("_(No message content)_")
        lines.append("")

        # Process attachments
        if email.attachments:
            for att in email.attachments:
                if att.is_eml:
                    continue  # Handled in nested emails section

                att_lines = self._process_attachment(att, email_number)
                if att_lines:
                    lines.extend(att_lines)
                    lines.append("")

        # Nested emails
        if email.nested_emails:
            lines.append("### Attached Emails")
            lines.append("")
            for nested_idx, nested_email in enumerate(email.nested_emails, 1):
                lines.append(f"#### Attached Email {nested_idx}: {nested_email.subject}")
                lines.append("")
                lines.append(f"- **From:** {nested_email.from_name} <{nested_email.from_addr}>")
                lines.append(f"- **Date:** {nested_email.date or 'Unknown'}")

                if nested_email.attachments:
                    lines.append(f"- **Attachments:** {len(nested_email.attachments)}")

                lines.append("")

                if nested_email.body_markdown:
                    lines.append(nested_email.body_markdown)
                else:
                    lines.append("_(No message content)_")
                lines.append("")

                # Process nested email attachments
                if nested_email.attachments:
                    for nested_att in nested_email.attachments:
                        if nested_att.is_eml:
                            continue

                        nested_att_lines = self._process_attachment(nested_att, email_number, nested_idx)
                        if nested_att_lines:
                            lines.extend(nested_att_lines)
                            lines.append("")

        return lines

    def _process_attachment(self, attachment: EmailAttachment, email_number: int,
                           nested_number: int = None) -> List[str]:
        """Process a single attachment - extract text or bundle for PDF"""
        lines = []

        # Try to extract text
        if self.extractor.can_extract_text(attachment.content_type, attachment.filename):
            extracted_text, success = self.extractor.extract_text(
                attachment.content,
                attachment.content_type,
                attachment.filename
            )

            if success and extracted_text:
                lines.append(f"### Attachment: {attachment.filename}")
                lines.append("")
                lines.append(f"**Type:** {attachment.content_type} | **Size:** {self._format_size(attachment.size)}")
                lines.append("")
                lines.append("**Extracted Content:**")
                lines.append("")
                lines.append("```")
                lines.append(extracted_text[:10000])  # Limit to 10KB of text
                if len(extracted_text) > 10000:
                    lines.append("\n... (content truncated)")
                lines.append("```")
                return lines

        # If text extraction failed or not applicable, check if it's an image
        if self.extractor.is_image(attachment.content_type, attachment.filename):
            self.images_for_bundle.append((attachment.content, attachment.filename))
            att_id = self._register_attachment(attachment, email_number, nested_number)
            lines.append(f"**Image Attachment:** `{attachment.filename}` ({self._format_size(attachment.size)}) → _Included in images-bundle.pdf as {att_id}_")
            return lines

        # Otherwise, it's a binary file
        self.binary_files_for_bundle.append((attachment.content, attachment.filename, attachment.content_type))
        att_id = self._register_attachment(attachment, email_number, nested_number)
        lines.append(f"**Binary Attachment:** `{attachment.filename}` ({attachment.content_type}, {self._format_size(attachment.size)}) → _{att_id} in binary-files directory_")
        return lines

    def _register_attachment(self, attachment: EmailAttachment, email_number: int, nested_number: int = None) -> str:
        """Register an attachment and return its ID"""
        self.attachment_counter += 1
        if nested_number:
            att_id = f"att-{email_number:04d}-{nested_number}-{self.attachment_counter:03d}"
        else:
            att_id = f"att-{email_number:04d}-{self.attachment_counter:03d}"

        self.attachment_registry[att_id] = attachment.filename
        return att_id

    def _format_size(self, size_bytes: int) -> str:
        """Format file size in human-readable format"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.1f} TB"

    def create_images_pdf(self):
        """Create a PDF bundle of all images"""
        if not self.images_for_bundle:
            logger.info("No images to bundle")
            return

        logger.info(f"Creating images bundle with {len(self.images_for_bundle)} images...")

        try:
            output_path = self.output_dir / "images-bundle.pdf"
            c = canvas.Canvas(str(output_path), pagesize=letter)
            page_width, page_height = letter

            for idx, (content, filename) in enumerate(self.images_for_bundle, 1):
                try:
                    # Load image
                    image = Image.open(io.BytesIO(content))

                    # Calculate scaling to fit page
                    img_width, img_height = image.size
                    scale = min((page_width - 100) / img_width, (page_height - 150) / img_height)

                    scaled_width = img_width * scale
                    scaled_height = img_height * scale

                    # Add page
                    c.setFont("Helvetica", 10)
                    c.drawString(50, page_height - 30, f"Image {idx}/{len(self.images_for_bundle)}: {filename}")

                    # Draw image
                    img_reader = ImageReader(io.BytesIO(content))
                    c.drawImage(img_reader, 50, page_height - 100 - scaled_height,
                               width=scaled_width, height=scaled_height)

                    c.showPage()

                except Exception as e:
                    logger.warning(f"Error adding image {filename} to PDF: {e}")

            c.save()
            logger.info(f"Saved images bundle to: {output_path}")

        except Exception as e:
            logger.error(f"Error creating images PDF: {e}")

    def save_binary_files(self):
        """Save binary files that couldn't be extracted"""
        if not self.binary_files_for_bundle:
            logger.info("No binary files to save")
            return

        logger.info(f"Saving {len(self.binary_files_for_bundle)} binary files...")

        binary_dir = self.output_dir / "binary-files"
        binary_dir.mkdir(parents=True, exist_ok=True)

        for content, filename, content_type in self.binary_files_for_bundle:
            try:
                # Find the attachment ID
                att_id = None
                for aid, fname in self.attachment_registry.items():
                    if fname == filename:
                        att_id = aid
                        break

                if att_id:
                    safe_filename = self._sanitize_filename(filename)
                    output_filename = f"{att_id}_{safe_filename}"
                    output_path = binary_dir / output_filename
                    output_path.write_bytes(content)
                    logger.debug(f"Saved: {output_filename}")

            except Exception as e:
                logger.error(f"Error saving {filename}: {e}")

        logger.info(f"Saved binary files to: {binary_dir}")

    def _sanitize_filename(self, filename: str) -> str:
        """Sanitize filename for safe filesystem storage"""
        sanitized = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', filename)
        sanitized = sanitized.strip('. ')
        if len(sanitized) > 200:
            parts = sanitized.rsplit('.', 1)
            if len(parts) == 2:
                sanitized = parts[0][:190] + '.' + parts[1]
            else:
                sanitized = sanitized[:200]
        return sanitized if sanitized else 'unnamed'

    def generate_manifest(self) -> Dict:
        """Generate manifest with statistics"""
        stats = self.extractor.get_stats()

        manifest = {
            'generated': datetime.now().isoformat(),
            'total_emails': len(self.emails),
            'extraction_stats': stats,
            'images_bundled': len(self.images_for_bundle),
            'binary_files_saved': len(self.binary_files_for_bundle),
            'emails': []
        }

        for idx, email in enumerate(self.emails, 1):
            email_info = {
                'number': idx,
                'subject': email.subject,
                'from': email.from_addr,
                'to': email.to_addrs,
                'date': email.date,
                'attachment_count': len(email.attachments),
                'has_nested_emails': len(email.nested_emails) > 0
            }
            manifest['emails'].append(email_info)

        return manifest

    def _create_summary(self):
        """Create summary file"""
        stats = self.extractor.get_stats()

        summary_lines = [
            "# Email Archive Summary (Enhanced)",
            "",
            f"**Total Emails:** {len(self.emails)}",
            f"**Date Range:** {self.emails[0].date if self.emails else 'N/A'} to {self.emails[-1].date if self.emails else 'N/A'}",
            "",
            "## Text Extraction Statistics",
            "",
            f"- **PDFs extracted:** {stats['pdf_extracted']}",
            f"- **DOCX files extracted:** {stats['docx_extracted']}",
            f"- **Text files extracted:** {stats['text_extracted']}",
            f"- **Extraction failures:** {stats['extraction_failed']}",
            f"- **Binary files kept:** {stats['binary_kept']}",
            "",
            "## Output Files",
            "",
            "- `emails.md` - All emails with extracted attachment text inline",
            f"- `images-bundle.pdf` - {len(self.images_for_bundle)} images combined",
            f"- `binary-files/` - {len(self.binary_files_for_bundle)} other attachments",
            "- `manifest.json` - Machine-readable metadata",
            "",
            "## How to Use with NotebookLM",
            "",
            "**Upload these files (3 total):**",
            "1. `emails.md` (contains all emails + extracted text)",
            "2. `images-bundle.pdf` (if you need image references)",
            "3. Selected files from `binary-files/` (if needed)",
            "",
            "Most content is now searchable in the main emails.md file!",
            ""
        ]

        summary_path = self.output_dir / "README.md"
        summary_path.write_text('\n'.join(summary_lines), encoding='utf-8')
        logger.info(f"Saved summary to: {summary_path}")

    def process(self):
        """Main processing workflow"""
        # Create output directory
        self.output_dir.mkdir(parents=True, exist_ok=True)

        # Process all emails
        self.process_all_emails()

        if not self.emails:
            logger.warning("No emails found to process")
            return

        # Generate markdown with extraction
        markdown_content = self.generate_markdown()
        markdown_path = self.output_dir / "emails.md"
        markdown_path.write_text(markdown_content, encoding='utf-8')
        logger.info(f"Saved markdown to: {markdown_path} ({len(markdown_content)} bytes)")

        # Create images PDF bundle
        self.create_images_pdf()

        # Save binary files
        self.save_binary_files()

        # Generate manifest
        manifest = self.generate_manifest()
        manifest_path = self.output_dir / "manifest.json"
        manifest_path.write_text(json.dumps(manifest, indent=2), encoding='utf-8')
        logger.info(f"Saved manifest to: {manifest_path}")

        # Create summary
        self._create_summary()

        # Print statistics
        stats = self.extractor.get_stats()
        logger.info("\n" + "="*60)
        logger.info("PROCESSING COMPLETE!")
        logger.info("="*60)
        logger.info(f"Total emails processed: {len(self.emails)}")
        logger.info(f"Text extracted from {stats['pdf_extracted']} PDFs")
        logger.info(f"Text extracted from {stats['docx_extracted']} DOCX files")
        logger.info(f"Images bundled: {len(self.images_for_bundle)}")
        logger.info(f"Binary files saved: {len(self.binary_files_for_bundle)}")
        logger.info("")
        logger.info("FILES TO UPLOAD TO NOTEBOOKLM:")
        logger.info(f"  1. {markdown_path}")
        if self.images_for_bundle:
            logger.info(f"  2. {self.output_dir / 'images-bundle.pdf'}")
        if self.binary_files_for_bundle:
            logger.info(f"  3. Selected files from {self.output_dir / 'binary-files/'}")
        logger.info("="*60)


def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description='Prepare email archives for LLM (Enhanced with text extraction)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Example usage:
  %(prog)s --input emails_comprehensive --output prepared-for-llm

  This will create:
  - emails.md (all emails + extracted text from PDFs/DOCX)
  - images-bundle.pdf (all images combined)
  - binary-files/ (other attachments)
        """
    )

    parser.add_argument('--input', '-i', required=True,
                       help='Input directory containing EML files')
    parser.add_argument('--output', '-o', default='prepared-for-llm-enhanced',
                       help='Output directory (default: prepared-for-llm-enhanced)')
    parser.add_argument('--no-nested', action='store_true',
                       help='Do not extract nested .eml attachments')
    parser.add_argument('--verbose', '-v', action='store_true',
                       help='Enable verbose logging')

    args = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    input_dir = Path(args.input)
    if not input_dir.exists():
        logger.error(f"Input directory does not exist: {input_dir}")
        return 1

    output_dir = Path(args.output)

    # Process emails
    processor = EnhancedEmailProcessor(
        input_dir=input_dir,
        output_dir=output_dir,
        extract_nested=not args.no_nested
    )

    try:
        processor.process()
        return 0
    except Exception as e:
        logger.error(f"Processing failed: {e}", exc_info=True)
        return 1


if __name__ == '__main__':
    import sys
    sys.exit(main())
