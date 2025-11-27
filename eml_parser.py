#!/usr/bin/env python3
"""
EML Parser - Extract structured data from EML email files
"""

import email
from email import policy
from email.parser import BytesParser
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import html2text
from bs4 import BeautifulSoup
import base64
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class EmailAttachment:
    """Represents an email attachment"""

    def __init__(self, filename: str, content: bytes, content_type: str, is_eml: bool = False):
        self.filename = filename
        self.content = content
        self.content_type = content_type
        self.is_eml = is_eml
        self.size = len(content)

    def __repr__(self):
        return f"EmailAttachment(filename='{self.filename}', type='{self.content_type}', size={self.size}, is_eml={self.is_eml})"


class ParsedEmail:
    """Represents a parsed email with all its components"""

    def __init__(self):
        self.message_id: Optional[str] = None
        self.subject: str = ""
        self.from_addr: str = ""
        self.from_name: str = ""
        self.to_addrs: List[str] = []
        self.cc_addrs: List[str] = []
        self.date: Optional[str] = None
        self.date_obj = None
        self.body_text: str = ""
        self.body_html: str = ""
        self.body_markdown: str = ""
        self.attachments: List[EmailAttachment] = []
        self.has_nested_emails: bool = False
        self.nested_emails: List['ParsedEmail'] = []
        self.original_filename: str = ""

    def __repr__(self):
        return f"ParsedEmail(subject='{self.subject[:50]}', from='{self.from_addr}', attachments={len(self.attachments)})"


class EMLParser:
    """Parser for EML email files"""

    def __init__(self, extract_nested_emails: bool = True):
        """
        Initialize the EML parser

        Args:
            extract_nested_emails: If True, recursively parse .eml attachments
        """
        self.extract_nested_emails = extract_nested_emails
        self.html_converter = html2text.HTML2Text()
        self.html_converter.ignore_links = False
        self.html_converter.ignore_images = False
        self.html_converter.ignore_emphasis = False
        self.html_converter.body_width = 0  # Don't wrap lines

    def parse_file(self, eml_path: Path) -> ParsedEmail:
        """
        Parse an EML file

        Args:
            eml_path: Path to the EML file

        Returns:
            ParsedEmail object
        """
        logger.debug(f"Parsing EML file: {eml_path}")

        with open(eml_path, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)

        parsed = ParsedEmail()
        parsed.original_filename = eml_path.name

        # Extract headers
        parsed.message_id = msg.get('Message-ID', '').strip('<>')
        parsed.subject = self._decode_header(msg.get('Subject', '(No Subject)'))

        # Parse From address
        from_header = msg.get('From', '')
        parsed.from_name, parsed.from_addr = parseaddr(from_header)
        if not parsed.from_name:
            parsed.from_name = parsed.from_addr

        # Parse To addresses
        to_header = msg.get('To', '')
        if to_header:
            parsed.to_addrs = [addr for _, addr in email.utils.getaddresses([to_header])]

        # Parse CC addresses
        cc_header = msg.get('Cc', '')
        if cc_header:
            parsed.cc_addrs = [addr for _, addr in email.utils.getaddresses([cc_header])]

        # Parse date
        date_header = msg.get('Date')
        if date_header:
            try:
                parsed.date_obj = parsedate_to_datetime(date_header)
                parsed.date = parsed.date_obj.strftime('%Y-%m-%d %H:%M:%S')
            except Exception as e:
                logger.warning(f"Could not parse date '{date_header}': {e}")
                parsed.date = date_header

        # Extract body
        parsed.body_text, parsed.body_html = self._extract_body(msg)

        # Convert HTML to Markdown if we have HTML body
        if parsed.body_html:
            parsed.body_markdown = self._html_to_markdown(parsed.body_html)
        elif parsed.body_text:
            parsed.body_markdown = parsed.body_text

        # Extract attachments
        parsed.attachments = self._extract_attachments(msg)

        # Check for nested emails and parse them if requested
        if self.extract_nested_emails:
            for attachment in parsed.attachments:
                if attachment.is_eml:
                    parsed.has_nested_emails = True
                    try:
                        # Save temporarily and parse
                        temp_path = Path(f"/tmp/{attachment.filename}")
                        temp_path.write_bytes(attachment.content)
                        nested_email = self.parse_file(temp_path)
                        parsed.nested_emails.append(nested_email)
                        temp_path.unlink()
                    except Exception as e:
                        logger.error(f"Error parsing nested email {attachment.filename}: {e}")

        return parsed

    def _decode_header(self, header: str) -> str:
        """Decode email header that might be encoded"""
        if not header:
            return ""

        decoded_parts = []
        for part, encoding in email.header.decode_header(header):
            if isinstance(part, bytes):
                try:
                    decoded_parts.append(part.decode(encoding or 'utf-8', errors='replace'))
                except:
                    decoded_parts.append(part.decode('utf-8', errors='replace'))
            else:
                decoded_parts.append(part)

        return ''.join(decoded_parts)

    def _extract_body(self, msg) -> Tuple[str, str]:
        """
        Extract both plain text and HTML body from email

        Returns:
            Tuple of (plain_text, html_text)
        """
        text_body = ""
        html_body = ""

        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get('Content-Disposition', ''))

                # Skip attachments
                if 'attachment' in content_disposition:
                    continue

                if content_type == 'text/plain' and not text_body:
                    try:
                        payload = part.get_payload(decode=True)
                        charset = part.get_content_charset() or 'utf-8'
                        text_body = payload.decode(charset, errors='replace')
                    except Exception as e:
                        logger.warning(f"Error decoding text/plain part: {e}")

                elif content_type == 'text/html' and not html_body:
                    try:
                        payload = part.get_payload(decode=True)
                        charset = part.get_content_charset() or 'utf-8'
                        html_body = payload.decode(charset, errors='replace')
                    except Exception as e:
                        logger.warning(f"Error decoding text/html part: {e}")
        else:
            # Not multipart - single body
            content_type = msg.get_content_type()
            try:
                payload = msg.get_payload(decode=True)
                charset = msg.get_content_charset() or 'utf-8'
                body_content = payload.decode(charset, errors='replace')

                if content_type == 'text/html':
                    html_body = body_content
                else:
                    text_body = body_content
            except Exception as e:
                logger.warning(f"Error decoding message body: {e}")

        return text_body.strip(), html_body.strip()

    def _html_to_markdown(self, html_content: str) -> str:
        """Convert HTML to clean Markdown"""
        try:
            # Clean up HTML first with BeautifulSoup
            soup = BeautifulSoup(html_content, 'html.parser')

            # Remove script and style elements
            for script in soup(["script", "style"]):
                script.decompose()

            # Convert to markdown
            markdown = self.html_converter.handle(str(soup))

            # Clean up excessive newlines
            while '\n\n\n' in markdown:
                markdown = markdown.replace('\n\n\n', '\n\n')

            return markdown.strip()

        except Exception as e:
            logger.warning(f"Error converting HTML to markdown: {e}")
            return html_content

    def _extract_attachments(self, msg) -> List[EmailAttachment]:
        """Extract all attachments from the email"""
        attachments = []

        for part in msg.walk():
            # Skip multipart containers
            if part.get_content_maintype() == 'multipart':
                continue

            # Skip text/plain and text/html parts that are the email body
            content_disposition = str(part.get('Content-Disposition', ''))
            if not content_disposition or 'attachment' not in content_disposition:
                # This might be the body, not an attachment
                if part.get_content_type() in ['text/plain', 'text/html']:
                    continue

            # Get filename
            filename = part.get_filename()
            if not filename:
                # Try to generate a filename
                ext = self._guess_extension(part.get_content_type())
                filename = f'attachment{len(attachments) + 1}{ext}'
            else:
                # Decode filename if needed
                filename = self._decode_header(filename)

            # Get content
            try:
                payload = part.get_payload(decode=True)
                if payload:
                    content_type = part.get_content_type()
                    is_eml = filename.lower().endswith('.eml') or content_type == 'message/rfc822'

                    attachment = EmailAttachment(
                        filename=filename,
                        content=payload,
                        content_type=content_type,
                        is_eml=is_eml
                    )
                    attachments.append(attachment)

            except Exception as e:
                logger.warning(f"Error extracting attachment '{filename}': {e}")

        return attachments

    def _guess_extension(self, content_type: str) -> str:
        """Guess file extension from content type"""
        extensions = {
            'text/plain': '.txt',
            'text/html': '.html',
            'application/pdf': '.pdf',
            'image/jpeg': '.jpg',
            'image/png': '.png',
            'image/gif': '.gif',
            'application/msword': '.doc',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document': '.docx',
            'application/vnd.ms-excel': '.xls',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx',
        }
        return extensions.get(content_type, '')


def test_parser():
    """Test the parser with a sample EML file"""
    import sys

    if len(sys.argv) < 2:
        print("Usage: python eml_parser.py <eml_file>")
        sys.exit(1)

    parser = EMLParser(extract_nested_emails=True)
    eml_path = Path(sys.argv[1])

    if not eml_path.exists():
        print(f"Error: File not found: {eml_path}")
        sys.exit(1)

    parsed = parser.parse_file(eml_path)

    print(f"Subject: {parsed.subject}")
    print(f"From: {parsed.from_name} <{parsed.from_addr}>")
    print(f"To: {', '.join(parsed.to_addrs)}")
    print(f"Date: {parsed.date}")
    print(f"Attachments: {len(parsed.attachments)}")
    for att in parsed.attachments:
        print(f"  - {att.filename} ({att.content_type}, {att.size} bytes, EML: {att.is_eml})")

    if parsed.nested_emails:
        print(f"\nNested Emails: {len(parsed.nested_emails)}")
        for nested in parsed.nested_emails:
            print(f"  - {nested.subject}")

    print(f"\nBody (first 500 chars):\n{parsed.body_markdown[:500]}")


if __name__ == '__main__':
    test_parser()
