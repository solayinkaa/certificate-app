"""
Certificate Distribution SaaS
=============================

A single-file Streamlit application that:

1. Takes a certificate template (.pdf, .png, .jpg/.jpeg) and lets the user
   calibrate where the recipient's Name and Certificate Number are printed,
   with a live rendered preview.
2. Reads a recipient list from an .xlsx file (columns: Name, CertNo, Email),
   sanitizes it, and applies a freemium row limit unless a Pro key is given.
3. Generates one flattened, permission-locked PDF per recipient entirely in
   memory and e-mails it over a single persistent SMTP connection, with
   live telemetry and a downloadable CSV report.

Deployment target: Streamlit Community Cloud (free Linux container) embedded
as an iframe on c2snet.org, with several concurrent users. For that reason:

* Nothing is ever written to disk. Every upload, intermediate document and
  generated certificate lives in RAM (bytes / io.BytesIO) and is scoped to the
  current Streamlit session, so concurrent users can never collide on files.
* Only PyMuPDF Base-14 fonts are used ("helv", "tiro", "cour"), so no .ttf
  file is required on the container.
* The dispatch loop runs on Streamlit's own script thread (no threading.Thread),
  which avoids "missing ScriptRunContext / NoSessionContext" crashes.
* Documents and buffers are closed after every row and gc.collect() runs every
  25 rows to keep RAM flat on the small free-tier container.
"""

# ---------------------------------------------------------------------------
# Standard library imports
# ---------------------------------------------------------------------------
import gc                      # Manual garbage collection for the anti-OOM policy
import hashlib                 # Stable short hash of the template (widget keys)
import hmac                    # Constant-time comparison of license keys
import io                      # In-memory byte streams (zero disk writes)
import os                      # Optional environment-variable key override
import re                      # Text sanitizing and e-mail validation
import secrets                 # Random owner password used to lock each PDF
import smtplib                 # SMTP client
import socket                  # Network error types for reconnect handling
import ssl                     # Verified TLS contexts for SMTP_SSL / STARTTLS
import time                    # Rate-limit delay between e-mails
from collections import deque  # Bounded telemetry log (constant memory)
from datetime import datetime  # Timestamps for log lines and report file name
from email.message import EmailMessage  # Modern MIME message builder
from email.utils import formatdate, make_msgid  # Standard e-mail headers

# ---------------------------------------------------------------------------
# Third-party imports
# ---------------------------------------------------------------------------
import pandas as pd            # Excel parsing / sanitizing (openpyxl engine)
import streamlit as st         # Web UI framework

# PyMuPDF is importable as "pymupdf" (current name) or "fitz" (legacy name).
# Trying both keeps the app working across PyMuPDF versions.
try:
    import pymupdf as fitz
except ImportError:  # pragma: no cover - only hit on very old PyMuPDF builds
    import fitz


# ---------------------------------------------------------------------------
# Page configuration. Must be the first Streamlit call in the script.
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Certificate Distribution",
    page_icon="🎓",
    layout="wide",
)


# ===========================================================================
# CONSTANTS
# ===========================================================================

# Pro license keys required by the product specification.
# NOTE: anything hard-coded here is visible to whoever can read the source
# (a public GitHub repo exposes its files). To rotate or hide keys
# without editing code, add more keys through the PRO_LICENSE_KEYS
# environment variable or st.secrets["PRO_LICENSE_KEYS"] (a TOML list or a
# comma-separated string, e.g. in Streamlit Community Cloud's Secrets box).
# Those are merged with this list at runtime.
VALID_KEYS = []  # Real keys live in Streamlit Secrets (PRO_LICENSE_KEYS)

FREE_ROW_LIMIT = 50                     # Max recipients on the free tier
REQUIRED_COLUMNS = ["Name", "CertNo", "Email"]
FONT_CHOICES = {                        # Base-14 fonts bundled in PyMuPDF
    "helv": "Helvetica (sans-serif)",
    "tiro": "Times Roman (serif)",
    "cour": "Courier (monospace)",
}
SAMPLE_NAME = "Jane Doe"                # Preview text for the Name field
SAMPLE_CERTNO = "CERT-001"              # Preview text for the CertNo field
SMTP_TIMEOUT = 20                       # Seconds, per the specification
GC_EVERY = 25                           # Run gc.collect() every N records
LOG_MAX_LINES = 200                     # Lines kept in the on-screen log
IMAGE_TEMPLATE_LONG_SIDE = 842.0        # Points (A4 long side) for image templates
PREVIEW_MAX_WIDTH_PX = 1100             # Preview raster width cap (RAM/CPU saver)
EMAIL_RE = re.compile(r"^[^@\s<>(),;:\"]+@[^@\s<>(),;:\"]+\.[A-Za-z]{2,}$")


# ===========================================================================
# LICENSE HELPERS
# ===========================================================================

def _extra_keys_from_config():
    """Collect optional extra Pro keys from st.secrets and the environment.

    Returns a list of non-empty, stripped key strings. Any error while reading
    st.secrets (for example when no secrets.toml exists) is ignored, because
    the hard-coded VALID_KEYS list is always available as the baseline.
    """
    raw_values = []
    try:
        if "PRO_LICENSE_KEYS" in st.secrets:
            secret = st.secrets["PRO_LICENSE_KEYS"]
            # Accept a TOML list (["A", "B"]) or a comma-separated string.
            if isinstance(secret, (list, tuple)):
                raw_values.extend(str(k) for k in secret)
            else:
                raw_values.append(str(secret))
    except Exception:
        pass
    raw_values.append(os.environ.get("PRO_LICENSE_KEYS", ""))
    keys = []
    for raw in raw_values:
        keys.extend(k.strip() for k in raw.split(",") if k.strip())
    return keys


def is_pro_key(candidate):
    """Return True if `candidate` matches any configured Pro key.

    hmac.compare_digest is used so the comparison time does not leak how many
    leading characters were correct.
    """
    candidate = (candidate or "").strip()
    if not candidate:
        return False
    for key in VALID_KEYS + _extra_keys_from_config():
        if hmac.compare_digest(candidate.encode("utf-8"), key.encode("utf-8")):
            return True
    return False


# ===========================================================================
# PDF / IMAGE ENGINE (PyMuPDF only, fully in memory)
# ===========================================================================

def hex_to_rgb01(hex_color):
    """Convert '#RRGGBB' (or '#RGB') into an (r, g, b) tuple of 0-1 floats."""
    value = (hex_color or "#000000").lstrip("#")
    if len(value) == 3:
        value = "".join(ch * 2 for ch in value)
    try:
        r, g, b = (int(value[i:i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        r, g, b = 0, 0, 0
    return (r / 255.0, g / 255.0, b / 255.0)


@st.cache_data(show_spinner=False, max_entries=16)
def normalize_template(file_bytes, extension):
    """Turn any supported upload into a clean single-page PDF byte string.

    * PDF: only the first page is kept and any /Rotate is removed so the
      coordinate system matches what the user sees in the preview.
    * PNG/JPG: the image is placed on a new page whose long side is
      IMAGE_TEMPLATE_LONG_SIDE points, preserving aspect ratio. This keeps
      font sizes meaningful no matter how many pixels the image has, while
      the image itself is embedded at full resolution.

    The result is cached by Streamlit (keyed on the bytes' content hash), so
    moving a slider never re-parses the upload. The cache is shared between
    sessions, which is safe because identical bytes give identical output.

    Returns (pdf_bytes, page_width, page_height).
    """
    extension = extension.lower().lstrip(".")
    if extension == "pdf":
        src = fitz.open(stream=file_bytes, filetype="pdf")
        try:
            if src.page_count < 1:
                raise ValueError("The PDF template has no pages.")
            if src.is_encrypted and not src.authenticate(""):
                raise ValueError("The PDF template is password protected.")
            if src.page_count > 1:
                src.select([0])
            page = src[0]
            if page.rotation and hasattr(page, "remove_rotation"):
                page.remove_rotation()
            pdf_bytes = src.tobytes(garbage=3, deflate=True)
        finally:
            src.close()
    else:
        img_doc = fitz.open(stream=file_bytes, filetype=extension)
        try:
            img_rect = img_doc[0].rect  # Image size as PyMuPDF reports it
        finally:
            img_doc.close()
        if img_rect.width <= 0 or img_rect.height <= 0:
            raise ValueError("The image template has invalid dimensions.")
        scale = IMAGE_TEMPLATE_LONG_SIDE / max(img_rect.width, img_rect.height)
        page_w, page_h = img_rect.width * scale, img_rect.height * scale
        out = fitz.open()
        try:
            page = out.new_page(width=page_w, height=page_h)
            page.insert_image(page.rect, stream=file_bytes)
            pdf_bytes = out.tobytes(garbage=3, deflate=True)
        finally:
            out.close()

    # Re-open once to report the final page size.
    check = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        rect = check[0].rect
        width, height = float(rect.width), float(rect.height)
    finally:
        check.close()
    return pdf_bytes, width, height


def fit_font_size(text, fontname, fontsize, max_width):
    """Shrink `fontsize` until `text` fits inside `max_width` points.

    A max_width of 0 (or less) disables auto-fit. The size never drops below
    6 pt so text always stays legible.
    """
    if max_width <= 0 or not text:
        return fontsize
    width = fitz.get_text_length(text, fontname=fontname, fontsize=fontsize)
    if width <= max_width:
        return fontsize
    return max(6.0, fontsize * max_width / width)


def draw_centered_text(page, text, center_x, baseline_y, fontname, fontsize,
                       color, max_width):
    """Write `text` horizontally centered on `center_x` at baseline `baseline_y`.

    The rendered width is measured with fitz.get_text_length for the chosen
    Base-14 font, so short and long names are both centered on the same
    axis. If `max_width` > 0 the font shrinks to keep long names inside that
    bounding width.
    """
    text = str(text or "")
    if not text:
        return
    size = fit_font_size(text, fontname, fontsize, max_width)
    width = fitz.get_text_length(text, fontname=fontname, fontsize=size)
    start_x = center_x - width / 2.0
    page.insert_text(
        fitz.Point(start_x, baseline_y),
        text,
        fontname=fontname,
        fontsize=size,
        color=color,
    )


def render_certificate(template_pdf, name, certno, layout):
    """Create one finished certificate and return it as an io.BytesIO.

    Steps: open the normalized template from bytes, stamp Name and CertNo,
    bake (flatten) annotations and form fields into page content, then save
    with AES-256 encryption that has no user password (opens anywhere) but a
    random owner password and no "modify" or "annotate" permission, so
    viewers treat the certificate as read-only.

    The fitz.Document is always closed before returning. The caller owns the
    returned BytesIO and must close it once the e-mail has been sent.
    """
    doc = fitz.open(stream=template_pdf, filetype="pdf")
    try:
        page = doc[0]
        draw_centered_text(page, name, layout["name_x"], layout["name_y"],
                           layout["font"], layout["name_size"],
                           layout["color"], layout["name_max_w"])
        draw_centered_text(page, certno, layout["cert_x"], layout["cert_y"],
                           layout["font"], layout["cert_size"],
                           layout["color"], layout["cert_max_w"])
        # Flatten any annotations / form widgets that came with the template.
        if hasattr(doc, "bake"):
            doc.bake()
        buffer = io.BytesIO()
        doc.save(
            buffer,
            garbage=3,
            deflate=True,
            encryption=fitz.PDF_ENCRYPT_AES_256,
            owner_pw=secrets.token_urlsafe(24),
            user_pw="",
            permissions=(fitz.PDF_PERM_PRINT | fitz.PDF_PERM_PRINT_HQ
                         | fitz.PDF_PERM_COPY | fitz.PDF_PERM_ACCESSIBILITY),
        )
        buffer.seek(0)
        return buffer
    finally:
        doc.close()


@st.cache_data(show_spinner=False, max_entries=64)
def render_preview_png(template_pdf, layout_items, show_guides):
    """Render the template with sample text into PNG bytes for the preview.

    `layout_items` is a tuple of (key, value) pairs so the arguments are
    hashable and Streamlit can cache identical previews. Optional red guide
    lines mark each X axis and baseline to help calibration. Rendering zoom
    is capped so the raster is at most PREVIEW_MAX_WIDTH_PX wide.
    """
    layout = dict(layout_items)
    doc = fitz.open(stream=template_pdf, filetype="pdf")
    try:
        page = doc[0]
        draw_centered_text(page, SAMPLE_NAME, layout["name_x"],
                           layout["name_y"], layout["font"],
                           layout["name_size"], layout["color"],
                           layout["name_max_w"])
        draw_centered_text(page, SAMPLE_CERTNO, layout["cert_x"],
                           layout["cert_y"], layout["font"],
                           layout["cert_size"], layout["color"],
                           layout["cert_max_w"])
        if show_guides:
            guide = (0.9, 0.1, 0.1)
            rect = page.rect
            for x, y, max_w in (
                (layout["name_x"], layout["name_y"], layout["name_max_w"]),
                (layout["cert_x"], layout["cert_y"], layout["cert_max_w"]),
            ):
                page.draw_line((x, rect.y0), (x, rect.y1), color=guide,
                               width=0.5, dashes="[3 3] 0")
                page.draw_line((rect.x0, y), (rect.x1, y), color=guide,
                               width=0.5, dashes="[3 3] 0")
                if max_w > 0:
                    page.draw_rect(fitz.Rect(x - max_w / 2, y - 4,
                                             x + max_w / 2, y + 4),
                                   color=guide, width=0.5)
        zoom = min(2.0, PREVIEW_MAX_WIDTH_PX / max(page.rect.width, 1.0))
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        png = pix.tobytes("png")
        del pix
        return png
    finally:
        doc.close()


# ===========================================================================
# EXCEL PARSING / SANITIZING
# ===========================================================================

def clean_certno(value):
    """Return CertNo as a clean string.

    Excel stores numbers as floats, so 1001 arrives as 1001.0. Integral
    floats become plain integers, any remaining trailing '.0' is removed,
    and NaN/None becomes an empty string.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    return re.sub(r"\.0+$", "", text) if re.fullmatch(r"-?\d+\.0+", text) else text


@st.cache_data(show_spinner=False, max_entries=16)
def load_recipients(xlsx_bytes):
    """Parse and sanitize the recipient spreadsheet.

    Returns (dataframe, missing_columns). When required columns are missing
    the dataframe is None and `missing_columns` lists what is absent.

    Sanitizing rules:
    * Headers are stripped and matched case-insensitively to Name / CertNo /
      Email (so " email " still works); they are renamed to canonical case.
    * All string cells are stripped; blank strings become missing values.
    * Completely empty rows are dropped.
    * CertNo is cast to a clean string without Excel's trailing ".0".
    """
    df = pd.read_excel(io.BytesIO(xlsx_bytes), engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    # Case-insensitive header matching onto canonical names.
    lookup = {c.lower(): c for c in df.columns}
    rename = {}
    for required in REQUIRED_COLUMNS:
        found = lookup.get(required.lower())
        if found is not None and found != required:
            rename[found] = required
    df = df.rename(columns=rename)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        return None, missing

    # Strip whitespace from every text cell and treat "" as missing.
    # (Checks each value rather than the column dtype, which differs between
    # pandas 2 "object" and pandas 3 "str" columns.)
    for col in df.columns:
        df[col] = df[col].map(lambda v: v.strip() if isinstance(v, str) else v)
    df = df.replace(r"^\s*$", pd.NA, regex=True)
    df = df.dropna(how="all").reset_index(drop=True)

    df["CertNo"] = df["CertNo"].map(clean_certno)
    df["Name"] = df["Name"].map(lambda v: "" if pd.isna(v) else str(v).strip())
    df["Email"] = df["Email"].map(lambda v: "" if pd.isna(v) else str(v).strip())
    return df, []


# ===========================================================================
# SMTP HELPERS
# ===========================================================================

def open_smtp(host, port, user, password):
    """Open and authenticate one SMTP connection.

    Port 465 uses implicit TLS (SMTP_SSL). Every other port uses a plain
    connection upgraded with STARTTLS. Both use a certificate-verifying TLS
    context and a 20 s timeout. Raises smtplib exceptions to the caller.
    """
    context = ssl.create_default_context()
    if int(port) == 465:
        server = smtplib.SMTP_SSL(host, int(port), timeout=SMTP_TIMEOUT,
                                  context=context)
        server.ehlo()
    else:
        server = smtplib.SMTP(host, int(port), timeout=SMTP_TIMEOUT)
        server.ehlo()
        server.starttls(context=context)
        server.ehlo()
    try:
        server.login(user, password)
    except Exception:
        close_smtp(server)
        raise
    return server


def close_smtp(server):
    """Close an SMTP connection politely, ignoring errors on a dead socket."""
    if server is None:
        return
    try:
        server.quit()
    except Exception:
        try:
            server.close()
        except Exception:
            pass


def fill_placeholders(template, name, certno):
    """Replace {Name} and {CertNo} in subject/body text.

    str.replace is used instead of str.format so any other braces the user
    types (e.g. in a signature) can never raise KeyError.
    """
    return (template or "").replace("{Name}", name).replace("{CertNo}", certno)


def safe_filename(text):
    """Make a file-name-safe token from CertNo (or a fallback)."""
    token = re.sub(r"[^A-Za-z0-9._-]+", "_", text or "").strip("._")
    return token[:80] or "certificate"


def build_message(sender, recipient, subject, body, pdf_bytes, filename):
    """Assemble the outgoing e-mail with the certificate attached."""
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=sender.split("@")[-1] or None)
    msg.set_content(body)
    msg.add_attachment(pdf_bytes, maintype="application", subtype="pdf",
                       filename=filename)
    return msg


# Errors that mean "the connection dropped", so reconnecting may help.
DISCONNECT_ERRORS = (smtplib.SMTPServerDisconnected, ConnectionError,
                     socket.timeout, ssl.SSLError)

AUTH_HELP = (
    "Authentication failed. The e-mail provider rejected the login. Either "
    "the App Password is wrong, or the provider blocked sign-in from this "
    "cloud server's IP address. For Gmail, enable 2-Step Verification and "
    "create an App Password; for cPanel, use the full mailbox address and "
    "its password, and check that the host allows external SMTP logins."
)


# ===========================================================================
# SESSION STATE (per browser tab, never shared between users)
# ===========================================================================

if "report_df" not in st.session_state:
    st.session_state.report_df = None        # Last batch report (DataFrame)
if "report_stamp" not in st.session_state:
    st.session_state.report_stamp = ""       # Timestamp used in file name


# ===========================================================================
# HEADER + FREEMIUM GATE
# ===========================================================================

st.title("🎓 Certificate Distribution")
st.caption("Design once, personalize every certificate, and e-mail them in one batch.")

license_key = st.text_input(
    "Pro License Key",
    type="password",
    placeholder="Leave empty to use the free tier (up to 50 recipients)",
    help="Enter your Pro key to unlock unlimited recipients.",
)
IS_PRO = is_pro_key(license_key)
if IS_PRO:
    st.success("Pro license active: unlimited recipients unlocked.")
elif license_key.strip():
    st.warning("That license key is not valid. The free tier limit of "
               f"{FREE_ROW_LIMIT} recipients applies.")


# ===========================================================================
# SIDEBAR: SMTP + E-MAIL CONFIGURATION
# ===========================================================================

with st.sidebar:
    st.header("✉️ E-mail settings")
    smtp_host = st.text_input("SMTP Server", value="smtp.gmail.com",
                              help="e.g. smtp.gmail.com or mail.yourdomain.com")
    smtp_port = st.number_input("SMTP Port", min_value=1, max_value=65535,
                                value=465, step=1,
                                help="465 = SSL, 587 (or any other) = STARTTLS")
    sender_email = st.text_input("Sender Email", placeholder="you@example.com")
    app_password = st.text_input("App Password", type="password")

    st.divider()
    email_subject = st.text_input("Email Subject",
                                  value="Your Certificate {CertNo}")
    email_body = st.text_area(
        "Email Body",
        height=180,
        value=("Dear {Name},\n\n"
               "Congratulations! Please find attached your certificate "
               "(No. {CertNo}).\n\nBest regards,\nThe C2S Team"),
        help="Use {Name} and {CertNo} as placeholders.",
    )
    send_delay = st.slider("Delay between e-mails (seconds)", 0.0, 10.0, 1.0,
                           0.1, help="Respects provider rate limits.")

    st.divider()
    if st.button("🔌 Test SMTP Connection", width="stretch"):
        if not (smtp_host.strip() and sender_email.strip() and app_password):
            st.error("Fill in server, sender e-mail and app password first.")
        else:
            test_server = None
            with st.spinner("Connecting..."):
                try:
                    test_server = open_smtp(smtp_host.strip(), smtp_port,
                                            sender_email.strip(), app_password)
                    st.success("Connection and login succeeded.")
                except smtplib.SMTPAuthenticationError:
                    st.error(AUTH_HELP)
                except Exception as exc:  # DNS, TLS, timeout, refused...
                    st.error(f"Connection failed: {type(exc).__name__}: {exc}")
                finally:
                    close_smtp(test_server)


# ===========================================================================
# STEP 1: TEMPLATE + LIVE CALIBRATION
# ===========================================================================

st.header("Step 1 · Template & coordinate calibration")
template_file = st.file_uploader("Certificate template",
                                 type=["pdf", "png", "jpg", "jpeg"])

template_pdf = None
layout = None
if template_file is not None:
    try:
        ext = template_file.name.rsplit(".", 1)[-1].lower()
        template_pdf, page_w, page_h = normalize_template(template_file.getvalue(), ext)
    except Exception as exc:
        st.error(f"Could not read the template: {exc}")
        template_pdf = None

if template_pdf is not None:
    # Slider keys include the template hash so the controls reset to sensible
    # defaults whenever a template with different dimensions is uploaded.
    tkey = hashlib.sha1(template_pdf).hexdigest()[:10]
    controls, preview = st.columns([1, 2], gap="large")

    with controls:
        font = st.selectbox("Font family", list(FONT_CHOICES),
                            format_func=lambda k: FONT_CHOICES[k],
                            key=f"font_{tkey}")
        hex_color = st.color_picker("Font color", "#1A1A1A",
                                    key=f"color_{tkey}")
        show_guides = st.checkbox("Show alignment guides", value=True,
                                  key=f"guides_{tkey}")
        st.caption(f"Page size: {page_w:.0f} × {page_h:.0f} pt. "
                   "X is the horizontal center of the text; Y is the "
                   "baseline measured from the top.")

        st.subheader("Name")
        name_x = st.slider("Name X (center)", 0.0, page_w, float(round(page_w / 2)), 1.0,
                           key=f"nx_{tkey}")
        name_y = st.slider("Name Y (baseline)", 0.0, page_h, float(round(page_h * 0.50)),
                           1.0, key=f"ny_{tkey}")
        name_size = st.slider("Name font size", 6.0, 120.0, 32.0, 1.0,
                              key=f"ns_{tkey}")
        name_max_w = st.slider("Name max width (0 = no auto-shrink)", 0.0,
                               page_w, float(round(page_w * 0.7)), 1.0, key=f"nw_{tkey}")

        st.subheader("Certificate No.")
        cert_x = st.slider("CertNo X (center)", 0.0, page_w, float(round(page_w / 2)), 1.0,
                           key=f"cx_{tkey}")
        cert_y = st.slider("CertNo Y (baseline)", 0.0, page_h, float(round(page_h * 0.65)),
                           1.0, key=f"cy_{tkey}")
        cert_size = st.slider("CertNo font size", 6.0, 72.0, 14.0, 1.0,
                              key=f"cs_{tkey}")
        cert_max_w = st.slider("CertNo max width (0 = no auto-shrink)", 0.0,
                               page_w, 0.0, 1.0, key=f"cw_{tkey}")

    layout = {
        "font": font,
        "color": hex_to_rgb01(hex_color),
        "name_x": name_x, "name_y": name_y,
        "name_size": name_size, "name_max_w": name_max_w,
        "cert_x": cert_x, "cert_y": cert_y,
        "cert_size": cert_size, "cert_max_w": cert_max_w,
    }

    with preview:
        try:
            png = render_preview_png(template_pdf, tuple(sorted(layout.items())),
                                     show_guides)
            st.image(png, caption="Live preview with sample data",
                     width="stretch")
        except Exception as exc:
            st.error(f"Preview failed: {exc}")
else:
    st.info("Upload a PDF, PNG or JPG template to start calibrating.")


# ===========================================================================
# STEP 2: RECIPIENT DATA
# ===========================================================================

st.header("Step 2 · Recipient data")
data_file = st.file_uploader("Recipient list (.xlsx with Name, CertNo, Email)",
                             type=["xlsx"])

recipients = None
if data_file is not None:
    try:
        full_df, missing_cols = load_recipients(data_file.getvalue())
    except Exception as exc:
        st.error(f"Could not read the Excel file: {exc}")
        st.stop()
    if missing_cols:
        st.error("The spreadsheet is missing required column(s): "
                 f"{', '.join(missing_cols)}. Expected headers: "
                 f"{', '.join(REQUIRED_COLUMNS)}.")
        st.stop()
    if full_df.empty:
        st.error("The spreadsheet has no data rows.")
        st.stop()

    total_rows = len(full_df)
    if not IS_PRO and total_rows > FREE_ROW_LIMIT:
        recipients = full_df.head(FREE_ROW_LIMIT).copy()
        st.info(f"🔒 Free tier: only the first {FREE_ROW_LIMIT} of "
                f"{total_rows} recipients will be processed. Enter a Pro "
                "License Key above to unlock unlimited recipients.")
    else:
        recipients = full_df.copy()

    bad_email_count = int((~recipients["Email"].map(
        lambda e: bool(EMAIL_RE.match(e)))).sum())
    m1, m2, m3 = st.columns(3)
    m1.metric("Rows in file", total_rows)
    m2.metric("Rows to process", len(recipients))
    m3.metric("Invalid e-mails", bad_email_count)
    if bad_email_count:
        st.warning(f"{bad_email_count} row(s) have a missing or malformed "
                   "e-mail address. They will be marked Failed in the report.")
    st.dataframe(recipients[REQUIRED_COLUMNS], width="stretch",
                 hide_index=True, height=300)
else:
    st.info("Upload an .xlsx file with the columns Name, CertNo and Email.")


# ===========================================================================
# STEP 3: BATCH DISPATCH + TELEMETRY
# ===========================================================================

st.header("Step 3 · Batch dispatch")

ready_problems = []
if template_pdf is None or layout is None:
    ready_problems.append("upload a template (Step 1)")
if recipients is None:
    ready_problems.append("upload the recipient list (Step 2)")
if not (smtp_host.strip() and sender_email.strip() and app_password):
    ready_problems.append("fill in the SMTP settings in the sidebar")
if ready_problems:
    st.info("Before dispatching, please " + ", ".join(ready_problems) + ".")

start = st.button("🚀 Start Bulk Dispatch", type="primary",
                  disabled=bool(ready_problems))

if start:
    host, user = smtp_host.strip(), sender_email.strip()
    total = len(recipients)
    progress = st.progress(0.0, text="Connecting to the mail server...")
    c1, c2, c3, c4 = st.columns(4)
    sent_box, failed_box, done_box, status_box = (c1.empty(), c2.empty(),
                                                  c3.empty(), c4.empty())
    log_box = st.empty()
    log_lines = deque(maxlen=LOG_MAX_LINES)
    report_rows = []
    sent = failed = 0

    def log(line):
        """Append a timestamped line and refresh the on-screen log."""
        log_lines.append(f"[{datetime.now():%H:%M:%S}] {line}")
        log_box.code("\n".join(log_lines), language=None)

    def refresh_metrics(done, state):
        """Redraw the counters (also keeps the WebSocket busy)."""
        sent_box.metric("Sent", sent)
        failed_box.metric("Failed", failed)
        done_box.metric("Processed", f"{done}/{total}")
        status_box.metric("Status", state)

    # ---- One persistent SMTP connection for the whole batch --------------
    server = None
    try:
        server = open_smtp(host, smtp_port, user, app_password)
        log(f"Connected to {host}:{int(smtp_port)} as {user}.")
    except smtplib.SMTPAuthenticationError:
        st.error(AUTH_HELP)
        st.stop()
    except Exception as exc:
        st.error(f"Could not connect to the mail server: "
                 f"{type(exc).__name__}: {exc}")
        st.stop()

    aborted_reason = ""
    try:
        refresh_metrics(0, "Running")
        for idx, row in enumerate(recipients.itertuples(index=False), start=1):
            name, certno, email = row.Name, row.CertNo, row.Email
            status, details = "Failed", ""
            pdf_buffer = pdf_bytes = msg = None
            try:
                if not EMAIL_RE.match(email or ""):
                    raise ValueError(f"Invalid e-mail address: '{email}'")
                if not name:
                    raise ValueError("Name is empty")

                # Build the certificate in memory.
                pdf_buffer = render_certificate(template_pdf, name, certno, layout)
                pdf_bytes = pdf_buffer.getvalue()
                msg = build_message(
                    user, email,
                    fill_placeholders(email_subject, name, certno),
                    fill_placeholders(email_body, name, certno),
                    pdf_bytes,
                    f"Certificate_{safe_filename(certno or name)}.pdf",
                )

                # Send, reconnecting exactly once if the server dropped us.
                try:
                    server.send_message(msg)
                except DISCONNECT_ERRORS as first_err:
                    log(f"Connection lost ({type(first_err).__name__}); "
                        "reconnecting once...")
                    close_smtp(server)
                    server = None
                    server = open_smtp(host, smtp_port, user, app_password)
                    server.send_message(msg)
                    details = "Sent after one reconnect"
                status = "Sent"
                sent += 1
                log(f"✅ {idx}/{total} sent to {email} ({certno})")
            except smtplib.SMTPAuthenticationError:
                # Only possible on reconnect: credentials or IP now blocked.
                failed += 1
                details = "Authentication rejected on reconnect"
                log(f"❌ {idx}/{total} {email}: {details}")
                aborted_reason = AUTH_HELP
            except smtplib.SMTPRecipientsRefused as exc:
                failed += 1
                details = "Recipient refused: " + "; ".join(
                    f"{code} {resp.decode(errors='replace') if isinstance(resp, bytes) else resp}"
                    for code, resp in exc.recipients.values())
                log(f"❌ {idx}/{total} {email}: {details}")
            except Exception as exc:
                failed += 1
                details = f"{type(exc).__name__}: {exc}"
                log(f"❌ {idx}/{total} {email}: {details}")
            finally:
                # Anti-OOM: release every per-row object immediately.
                if pdf_buffer is not None:
                    pdf_buffer.close()
                del pdf_buffer, pdf_bytes, msg
                report_rows.append({"Email": email, "CertNo": certno,
                                    "Status": status, "Error_Details": details})

            progress.progress(idx / total, text=f"Processed {idx} of {total}")
            refresh_metrics(idx, "Running")
            if idx % GC_EVERY == 0:
                gc.collect()

            # A dead connection that could not be re-opened stops the batch.
            if aborted_reason or server is None:
                aborted_reason = aborted_reason or "Mail server unavailable"
                for rest in recipients.iloc[idx:].itertuples(index=False):
                    report_rows.append({"Email": rest.Email,
                                        "CertNo": rest.CertNo,
                                        "Status": "Skipped",
                                        "Error_Details": "Batch aborted: "
                                                         "mail server unavailable"})
                break

            # Rate-limit delay, sliced so the UI updates while waiting.
            if idx < total and send_delay > 0:
                deadline = time.monotonic() + send_delay
                while True:
                    remaining = deadline - time.monotonic()
                    if remaining <= 0:
                        break
                    status_box.metric("Status", f"Waiting {remaining:.1f}s")
                    time.sleep(min(0.25, remaining))
    finally:
        close_smtp(server)
        gc.collect()

    if aborted_reason:
        refresh_metrics(len(report_rows), "Aborted")
        st.error(aborted_reason)
    else:
        refresh_metrics(total, "Finished")
        progress.progress(1.0, text="Batch complete")
        st.success(f"Batch complete: {sent} sent, {failed} failed.")

    # Keep the report in session state so the download button survives the
    # rerun that Streamlit performs when the button is clicked.
    st.session_state.report_df = pd.DataFrame(
        report_rows, columns=["Email", "CertNo", "Status", "Error_Details"])
    st.session_state.report_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

if st.session_state.report_df is not None:
    report = st.session_state.report_df
    st.subheader("Summary report")
    st.dataframe(report, width="stretch", hide_index=True, height=260)
    st.download_button(
        "⬇️ Download CSV report",
        data=report.to_csv(index=False).encode("utf-8"),
        file_name=f"dispatch_report_{st.session_state.report_stamp}.csv",
        mime="text/csv",
    )
