#!/usr/bin/env python3
"""
Prepare Emails for LLM - Convert email archives to LLM-friendly format

This script processes a directory of EML files and converts them into a format
suitable for ingestion into LLMs like NotebookLM.
"""

import argparse
import json
import logging
import shutil
from pathlib import Path
from typing import List, Dict
from datetime import datetime
from eml_parser import EMLParser, ParsedEmail, EmailAttachment
import re

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class EmailProcessor:
    """Process emails and generate LLM-friendly output"""

    def __init__(self, input_dir: Path, output_dir: Path, extract_nested: bool = True):
        """
        Initialize the email processor

        Args:
            input_dir: Directory containing EML files
            output_dir: Directory for output files
            extract_nested: Whether to recursively extract nested EML attachments
        """
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.parser = EMLParser(extract_nested_emails=extract_nested)
        self.emails: List[ParsedEmail] = []
        self.attachment_counter = 0
        self.attachment_registry: Dict[str, str] = {}  # Maps attachment ID to filename

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
        """Generate a single markdown file with all emails"""
        logger.info("Generating markdown...")

        lines = []

        # Header
        lines.append("# Email Archive")
        lines.append("")
        lines.append(f"**Total Emails:** {len(self.emails)}")
        lines.append(f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append("")
        lines.append("---")
        lines.append("")

        # Process each email
        for idx, email in enumerate(self.emails, 1):
            lines.extend(self._format_email_as_markdown(email, idx))
            lines.append("")
            lines.append("---")
            lines.append("")

        return '\n'.join(lines)

    def _format_email_as_markdown(self, email: ParsedEmail, email_number: int) -> List[str]:
        """Format a single email as markdown"""
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

        # Attachments
        if email.attachments:
            lines.append(f"- **Attachments:** {len(email.attachments)}")
            for att in email.attachments:
                att_id = self._register_attachment(att, email_number)
                if att.is_eml:
                    lines.append(f"  - `{att.filename}` (email message, parsed below)")
                else:
                    lines.append(f"  - `{att.filename}` ({att.content_type}, {self._format_size(att.size)}) [→ {att_id}]")

        lines.append("")

        # Body
        lines.append("**Message:**")
        lines.append("")
        if email.body_markdown:
            lines.append(email.body_markdown)
        else:
            lines.append("_(No message content)_")
        lines.append("")

        # Nested emails (from .eml attachments)
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
                    for nested_att in nested_email.attachments:
                        nested_att_id = self._register_attachment(nested_att, email_number, nested_idx)
                        lines.append(f"  - `{nested_att.filename}` ({nested_att.content_type}, {self._format_size(nested_att.size)}) [→ {nested_att_id}]")
                lines.append("")
                if nested_email.body_markdown:
                    lines.append(nested_email.body_markdown)
                else:
                    lines.append("_(No message content)_")
                lines.append("")

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

    def save_attachments(self):
        """Save all non-EML attachments to the output directory"""
        logger.info("Saving attachments...")

        attachments_dir = self.output_dir / "attachments"
        attachments_dir.mkdir(parents=True, exist_ok=True)

        attachment_count = 0

        for email_idx, email in enumerate(self.emails, 1):
            # Save direct attachments
            for att in email.attachments:
                if not att.is_eml:
                    self._save_attachment(att, attachments_dir, email_idx)
                    attachment_count += 1

            # Save attachments from nested emails
            for nested_idx, nested_email in enumerate(email.nested_emails, 1):
                for nested_att in nested_email.attachments:
                    if not nested_att.is_eml:
                        self._save_attachment(nested_att, attachments_dir, email_idx, nested_idx)
                        attachment_count += 1

        logger.info(f"Saved {attachment_count} attachments")

    def _save_attachment(self, attachment: EmailAttachment, output_dir: Path,
                        email_number: int, nested_number: int = None):
        """Save a single attachment to disk"""
        try:
            # Find the attachment ID
            att_id = None
            for aid, filename in self.attachment_registry.items():
                if filename == attachment.filename and aid.startswith(f"att-{email_number:04d}"):
                    if nested_number and f"-{nested_number}-" in aid:
                        att_id = aid
                        break
                    elif not nested_number and f"-{nested_number}-" not in aid:
                        att_id = aid
                        break

            if not att_id:
                logger.warning(f"Could not find attachment ID for {attachment.filename}")
                return

            # Sanitize filename
            safe_filename = self._sanitize_filename(attachment.filename)
            output_filename = f"{att_id}_{safe_filename}"
            output_path = output_dir / output_filename

            # Save the file
            output_path.write_bytes(attachment.content)
            logger.debug(f"Saved attachment: {output_filename}")

        except Exception as e:
            logger.error(f"Error saving attachment {attachment.filename}: {e}")

    def _sanitize_filename(self, filename: str) -> str:
        """Sanitize filename for safe filesystem storage"""
        # Remove or replace invalid characters
        sanitized = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', filename)
        sanitized = sanitized.strip('. ')
        if len(sanitized) > 200:
            # Keep extension
            parts = sanitized.rsplit('.', 1)
            if len(parts) == 2:
                sanitized = parts[0][:190] + '.' + parts[1]
            else:
                sanitized = sanitized[:200]
        return sanitized if sanitized else 'unnamed'

    def generate_manifest(self) -> Dict:
        """Generate a manifest with metadata about the processed emails"""
        manifest = {
            'generated': datetime.now().isoformat(),
            'total_emails': len(self.emails),
            'total_attachments': len(self.attachment_registry),
            'emails': []
        }

        for idx, email in enumerate(self.emails, 1):
            email_info = {
                'number': idx,
                'subject': email.subject,
                'from': email.from_addr,
                'to': email.to_addrs,
                'date': email.date,
                'message_id': email.message_id,
                'has_attachments': len(email.attachments) > 0,
                'attachment_count': len(email.attachments),
                'has_nested_emails': len(email.nested_emails) > 0,
                'nested_email_count': len(email.nested_emails)
            }
            manifest['emails'].append(email_info)

        return manifest

    def process(self):
        """Main processing workflow"""
        # Create output directory
        self.output_dir.mkdir(parents=True, exist_ok=True)

        # Process all emails
        self.process_all_emails()

        if not self.emails:
            logger.warning("No emails found to process")
            return

        # Generate markdown
        markdown_content = self.generate_markdown()
        markdown_path = self.output_dir / "emails.md"
        markdown_path.write_text(markdown_content, encoding='utf-8')
        logger.info(f"Saved markdown to: {markdown_path}")

        # Save attachments
        self.save_attachments()

        # Generate manifest
        manifest = self.generate_manifest()
        manifest_path = self.output_dir / "manifest.json"
        manifest_path.write_text(json.dumps(manifest, indent=2), encoding='utf-8')
        logger.info(f"Saved manifest to: {manifest_path}")

        # Create a summary file
        self._create_summary()

        logger.info("Processing complete!")
        logger.info(f"Output directory: {self.output_dir}")

    def _create_summary(self):
        """Create a summary/index file"""
        summary_lines = [
            "# Email Archive Summary",
            "",
            f"**Total Emails:** {len(self.emails)}",
            f"**Total Attachments:** {len(self.attachment_registry)}",
            f"**Date Range:** {self.emails[0].date if self.emails else 'N/A'} to {self.emails[-1].date if self.emails else 'N/A'}",
            "",
            "## Files in This Archive",
            "",
            "- `emails.md` - All emails in chronological order (main file)",
            "- `manifest.json` - Machine-readable metadata",
            "- `attachments/` - All email attachments (referenced in emails.md)",
            "",
            "## How to Use with NotebookLM",
            "",
            "1. Upload `emails.md` to NotebookLM",
            "2. Upload select files from the `attachments/` directory (up to 49 more files)",
            "3. Ask NotebookLM questions about the emails",
            "",
            "## Email List",
            ""
        ]

        for idx, email in enumerate(self.emails, 1):
            summary_lines.append(f"{idx}. **{email.subject}** - {email.from_addr} ({email.date})")

        summary_path = self.output_dir / "README.md"
        summary_path.write_text('\n'.join(summary_lines), encoding='utf-8')
        logger.info(f"Saved summary to: {summary_path}")


def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description='Prepare email archives for LLM ingestion',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Example usage:
  %(prog)s --input emails_comprehensive --output prepared-for-llm

  This will create:
  - prepared-for-llm/emails.md (all emails in one markdown file)
  - prepared-for-llm/attachments/ (all non-EML attachments)
  - prepared-for-llm/manifest.json (metadata)
  - prepared-for-llm/README.md (summary)
        """
    )

    parser.add_argument('--input', '-i', required=True,
                       help='Input directory containing EML files')
    parser.add_argument('--output', '-o', default='prepared-for-llm',
                       help='Output directory (default: prepared-for-llm)')
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
    processor = EmailProcessor(
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
