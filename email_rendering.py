import html
import re
from collections import OrderedDict
from datetime import datetime

try:
    from bs4 import BeautifulSoup, Comment
except ImportError:
    BeautifulSoup = None
    Comment = ()


BODY_FORMAT_GMAIL_PDF = "gmail_pdf"
BODY_FORMAT_ORIGINAL_HTML = "original_html"
BODY_FORMAT_BOTH = "both"

BODY_FORMAT_LABELS = OrderedDict([
    (BODY_FORMAT_GMAIL_PDF, "PDF estilo Gmail"),
    (BODY_FORMAT_ORIGINAL_HTML, "HTML original"),
    (BODY_FORMAT_BOTH, "Ambos"),
])
BODY_FORMAT_LABEL_VALUES = list(BODY_FORMAT_LABELS.values())
BODY_FORMAT_LABEL_TO_KEY = {label: key for key, label in BODY_FORMAT_LABELS.items()}

SAFE_STYLE_PROPERTIES = {
    "background-color",
    "border",
    "border-bottom",
    "border-collapse",
    "border-left",
    "border-right",
    "border-spacing",
    "border-top",
    "color",
    "display",
    "font-family",
    "font-size",
    "font-style",
    "font-weight",
    "line-height",
    "text-align",
    "text-decoration",
    "vertical-align",
    "white-space",
}
BLOCKED_STYLE_PREFIXES = (
    "bottom",
    "break-",
    "clear",
    "float",
    "height",
    "left",
    "margin",
    "max-",
    "min-",
    "mso-",
    "padding",
    "page-break",
    "position",
    "right",
    "top",
    "width",
)
DROP_TAGS = {"head", "meta", "link", "style", "script", "title", "base"}
UNWRAP_TAGS = {"html", "body", "tbody", "thead", "tfoot"}
ALLOWED_ALIGN = {"left", "center", "right", "justify"}
SAFE_LINK_SCHEMES = ("http://", "https://", "mailto:", "#", "data:")


def normalize_body_format(value):
    if value in BODY_FORMAT_LABELS:
        return value
    return BODY_FORMAT_LABEL_TO_KEY.get(value, BODY_FORMAT_GMAIL_PDF)


def body_format_label(value):
    return BODY_FORMAT_LABELS[normalize_body_format(value)]


def body_format_needs_pdf(value):
    return normalize_body_format(value) in {BODY_FORMAT_GMAIL_PDF, BODY_FORMAT_BOTH}


def body_format_needs_original_html(value):
    return normalize_body_format(value) in {BODY_FORMAT_ORIGINAL_HTML, BODY_FORMAT_BOTH}


def plain_text_to_html_fragment(text):
    return f"<pre>{html.escape(text or '')}</pre>"


def build_original_email_document(source_html):
    content = source_html or ""
    if re.search(r"<html\b", content, flags=re.IGNORECASE):
        return content
    return (
        "<html><head><meta charset=\"utf-8\"></head>"
        f"<body>{content}</body></html>"
    )


def sanitize_email_html(source_html):
    if not source_html:
        return ""
    if BeautifulSoup is None:
        return _sanitize_email_html_fallback(source_html)

    soup = BeautifulSoup(source_html, "html.parser")
    root = soup.body or soup

    for comment in root.find_all(string=lambda value: isinstance(value, Comment)):
        comment.extract()

    for tag in list(root.find_all(True)):
        if tag.attrs is None:
            continue
        name = (tag.name or "").lower()
        if name in DROP_TAGS or name == "xml" or ":" in name:
            tag.decompose()
            continue
        if name in UNWRAP_TAGS:
            tag.unwrap()
            continue
        _sanitize_tag(tag)

    if getattr(root, "attrs", None):
        root.attrs = {}

    cleaned = "".join(str(node) for node in root.contents).strip()
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    return cleaned


def build_gmail_print_document(
    account_email,
    subject,
    sender,
    recipient,
    cc,
    sent_at,
    body_fragment,
    print_timestamp=None,
):
    safe_account = html.escape(account_email or "")
    safe_subject = html.escape(subject or "")
    safe_sender = html.escape(sender or "")
    safe_recipient = html.escape(recipient or "")
    safe_cc = html.escape(cc or "")
    safe_sent_at = html.escape(sent_at or "")
    safe_print_timestamp = html.escape(
        print_timestamp or _format_print_timestamp(datetime.now())
    )
    safe_title = html.escape(f"{account_email} Mail - {subject}".strip(" -"))
    cc_row = f"<tr><td class=\"lbl\">Cc:</td><td>{safe_cc}</td></tr>" if safe_cc else ""

    return f"""<html>
<head>
<meta charset="utf-8">
<style>
    @page {{ size: Letter; margin: 12mm 10mm 16mm 10mm; }}
    * {{ box-sizing: border-box; }}
    body {{
        margin: 0;
        color: #202124;
        font-family: Arial, Helvetica, sans-serif;
        font-size: 10pt;
        line-height: 1.45;
    }}
    .print-meta {{
        display: flex;
        align-items: center;
        justify-content: space-between;
        color: #5f6368;
        font-size: 8.5pt;
        margin-bottom: 12px;
    }}
    .print-meta .center {{
        flex: 1;
        text-align: center;
        padding: 0 12px;
    }}
    .top-bar {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-bottom: 1px solid #dadce0;
        padding-bottom: 8px;
        margin-bottom: 10px;
    }}
    .gmail-logo {{
        display: flex;
        align-items: center;
        gap: 8px;
    }}
    .gmail-text {{
        color: #5f6368;
        font-size: 18pt;
        font-weight: normal;
    }}
    .account {{
        color: #202124;
        font-size: 9pt;
        font-weight: bold;
    }}
    .subject {{
        font-size: 16pt;
        font-weight: bold;
        margin: 0 0 2px 0;
    }}
    .msg-count {{
        color: #5f6368;
        font-size: 9pt;
        margin-bottom: 8px;
    }}
    .msg-info {{
        border-top: 1px solid #dadce0;
        border-bottom: 1px solid #dadce0;
        padding: 8px 0;
        margin-bottom: 16px;
    }}
    .msg-top {{
        display: flex;
        justify-content: space-between;
        gap: 12px;
        margin-bottom: 2px;
    }}
    .msg-from {{
        color: #202124;
        font-size: 10pt;
        font-weight: bold;
    }}
    .msg-date {{
        color: #5f6368;
        font-size: 9pt;
        white-space: nowrap;
    }}
    .msg-details {{
        color: #5f6368;
        font-size: 9pt;
    }}
    .msg-details table {{
        border-collapse: collapse;
    }}
    .msg-details td {{
        padding: 0 4px 0 0;
        vertical-align: top;
    }}
    .msg-details .lbl {{
        white-space: nowrap;
    }}
    .message-body {{
        color: #202124;
        font-size: 10pt;
        line-height: 1.45;
        overflow-wrap: break-word;
        word-break: break-word;
    }}
    .message-body p {{
        margin: 0 0 12px 0;
    }}
    .message-body img {{
        max-width: 100%;
        height: auto !important;
    }}
    .message-body table {{
        max-width: 100% !important;
        border-collapse: separate;
    }}
    .message-body td,
    .message-body th {{
        vertical-align: top;
    }}
    .message-body blockquote {{
        margin: 6px 0 6px 8px;
        padding-left: 12px;
        border-left: 2px solid #dadce0;
        color: #5f6368;
    }}
    .message-body pre {{
        white-space: pre-wrap;
        font-family: Consolas, "Courier New", monospace;
        margin: 0;
    }}
</style>
</head>
<body>
    <div class="print-meta">
        <span>{safe_print_timestamp}</span>
        <span class="center">{safe_title}</span>
        <span></span>
    </div>
    <div class="top-bar">
        <div class="gmail-logo">
            <svg width="34" height="26" viewBox="0 0 75 56" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
                <path fill="#ea4335" d="M67.6 12.2 37.5 34.9 7.4 12.2A4.3 4.3 0 0 0 0 15.6v24.8c0 2.4 1.9 4.3 4.3 4.3h8.6V22.4l24.6 18.5 24.6-18.5v22.3h8.6c2.4 0 4.3-1.9 4.3-4.3V15.6a4.3 4.3 0 0 0-7.4-3.4Z"/>
                <path fill="#34a853" d="M62.1 44.7V22.4l12.9-9.6v27.6c0 2.4-1.9 4.3-4.3 4.3h-8.6Z"/>
                <path fill="#4285f4" d="M0 12.8v27.6c0 2.4 1.9 4.3 4.3 4.3h8.6V22.4L0 12.8Z"/>
                <path fill="#fbbc04" d="M62.1 22.4v22.3H12.9V22.4l24.6 18.5 24.6-18.5Z"/>
            </svg>
            <span class="gmail-text">Gmail</span>
        </div>
        <span class="account">{safe_account}</span>
    </div>
    <div class="subject">{safe_subject}</div>
    <div class="msg-count">1 mensaje</div>
    <div class="msg-info">
        <div class="msg-top">
            <span class="msg-from">{safe_sender}</span>
            <span class="msg-date">{safe_sent_at}</span>
        </div>
        <div class="msg-details">
            <table>
                <tr><td class="lbl">To:</td><td>{safe_recipient}</td></tr>
                {cc_row}
            </table>
        </div>
    </div>
    <div class="message-body">{body_fragment}</div>
</body>
</html>"""


def _sanitize_tag(tag):
    derived_styles = []
    align_value = str(tag.attrs.get("align", "")).strip().lower()
    if align_value in ALLOWED_ALIGN:
        derived_styles.append(f"text-align:{align_value}")

    font_color = str(tag.attrs.get("color", "")).strip()
    if font_color:
        derived_styles.append(f"color:{font_color}")

    bgcolor = str(tag.attrs.get("bgcolor", "")).strip()
    if bgcolor:
        derived_styles.append(f"background-color:{bgcolor}")

    style = _sanitize_inline_style(str(tag.attrs.get("style", "")))
    if derived_styles:
        merged = ";".join(derived_styles + ([style] if style else []))
        style = _sanitize_inline_style(merged)

    cleaned = {}
    if tag.name == "a":
        href = str(tag.attrs.get("href", "")).strip()
        if _is_safe_link(href):
            cleaned["href"] = href
        title = str(tag.attrs.get("title", "")).strip()
        if title:
            cleaned["title"] = title
        target = str(tag.attrs.get("target", "")).strip()
        if target in {"_blank", "_self"}:
            cleaned["target"] = target
    elif tag.name == "img":
        src = str(tag.attrs.get("src", "")).strip()
        if src:
            cleaned["src"] = src
        alt = str(tag.attrs.get("alt", "")).strip()
        if alt:
            cleaned["alt"] = alt
        title = str(tag.attrs.get("title", "")).strip()
        if title:
            cleaned["title"] = title
        for attr in ("width", "height"):
            value = str(tag.attrs.get(attr, "")).strip()
            if _is_safe_dimension(value):
                cleaned[attr] = value
    else:
        for attr in ("colspan", "rowspan"):
            value = str(tag.attrs.get(attr, "")).strip()
            if value.isdigit():
                cleaned[attr] = value

    if style:
        cleaned["style"] = style

    tag.attrs = cleaned


def _sanitize_inline_style(style):
    safe_parts = []
    for rule in style.split(";"):
        if ":" not in rule:
            continue
        prop, value = rule.split(":", 1)
        prop = prop.strip().lower()
        value = value.strip()
        if not prop or not value:
            continue
        if prop not in SAFE_STYLE_PROPERTIES:
            if any(prop.startswith(prefix) for prefix in BLOCKED_STYLE_PREFIXES):
                continue
            continue
        if any(prop.startswith(prefix) for prefix in BLOCKED_STYLE_PREFIXES):
            continue
        if not _is_safe_style_value(value):
            continue
        if prop == "font-size" and not _is_safe_font_size(value):
            continue
        safe_parts.append(f"{prop}:{value}")
    return ";".join(safe_parts)


def _is_safe_style_value(value):
    lowered = value.lower()
    if any(token in lowered for token in ("expression(", "javascript:", "vbscript:", "url(")):
        return False
    return True


def _is_safe_font_size(value):
    match = re.match(r"^\s*(\d+(?:\.\d+)?)\s*(px|pt|em|rem|%)?\s*$", value, flags=re.IGNORECASE)
    if not match:
        return False
    size = float(match.group(1))
    unit = (match.group(2) or "px").lower()
    limits = {
        "px": 42,
        "pt": 32,
        "em": 3,
        "rem": 3,
        "%": 250,
    }
    return 0 < size <= limits[unit]


def _is_safe_link(value):
    if not value:
        return False
    lowered = value.lower()
    return lowered.startswith(SAFE_LINK_SCHEMES)


def _is_safe_dimension(value):
    return bool(re.match(r"^\d{1,4}%?$", value))


def _format_print_timestamp(value):
    return (
        value.strftime("%m/%d/%y, %I:%M %p")
        .replace("/0", "/")
        .lstrip("0")
        .lower()
    )


def _sanitize_email_html_fallback(source_html):
    cleaned = re.sub(r"<!--.*?-->", "", source_html, flags=re.DOTALL)
    cleaned = re.sub(r"<!DOCTYPE.*?>", "", cleaned, flags=re.IGNORECASE | re.DOTALL)
    cleaned = re.sub(
        r"<(script|style|meta|link|title|base)\b.*?>.*?</\1>",
        "",
        cleaned,
        flags=re.IGNORECASE | re.DOTALL,
    )
    cleaned = re.sub(r"<head\b.*?>.*?</head>", "", cleaned, flags=re.IGNORECASE | re.DOTALL)
    body_match = re.search(r"<body\b[^>]*>(.*)</body>", cleaned, flags=re.IGNORECASE | re.DOTALL)
    if body_match:
        cleaned = body_match.group(1)
    cleaned = re.sub(r"</?(html|body|tbody|thead|tfoot)\b[^>]*>", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"</?\w+:\w+[^>]*>", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s(on\w+|class|id|lang|face)=([\"']).*?\2", "", cleaned, flags=re.IGNORECASE | re.DOTALL)
    cleaned = re.sub(
        r"\sstyle=([\"'])(.*?)\1",
        lambda match: _replace_style_attr(match.group(2)),
        cleaned,
        flags=re.IGNORECASE | re.DOTALL,
    )
    return cleaned.strip()


def _replace_style_attr(style_text):
    style = _sanitize_inline_style(style_text)
    if not style:
        return ""
    return f' style="{style}"'
