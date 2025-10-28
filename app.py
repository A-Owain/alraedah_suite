# ============================================================
# Alraedah Suite — Minimal Font-Safe Signature Packager (Demo)
# ============================================================

import os
import io
import zipfile
import pathlib
from datetime import datetime

import streamlit as st
import pandas as pd

# ReportLab
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Timezone (optional; good practice for stamps)
from zoneinfo import ZoneInfo
LOCAL_TZ = ZoneInfo("Asia/Riyadh")

# -----------------------
# App Config
# -----------------------
st.set_page_config(page_title="Alraedah Suite — Signature Packager (Font Safe)", layout="wide")
st.title("Alraedah Suite — Font-Safe Signature PDF Packager (Demo)")

# -----------------------
# Font Registration
# -----------------------
@st.cache_resource
def register_fonts_once():
    """
    Register custom fonts for ReportLab safely on Streamlit Cloud.
    Looks for assets/fonts/PingARLT-Regular.ttf and PingARLT-Bold.ttf.
    Falls back automatically if not found.
    """
    base_dir = pathlib.Path(__file__).parent
    fonts_dir = base_dir / "assets" / "fonts"

    font_map = {
        "PingRegular": "PingARLT-Regular.ttf",
        "PingBold": "PingARLT-Bold.ttf",
    }

    for name, filename in font_map.items():
        ttf_path = fonts_dir / filename
        try:
            if ttf_path.exists():
                pdfmetrics.registerFont(TTFont(name, str(ttf_path)))
            else:
                st.warning(f"[Font] Missing file: {ttf_path} — will fallback for '{name}'.")
        except Exception as e:
            st.warning(f"[Font] Could not register '{name}': {e}")

    # optional family alias
    try:
        pdfmetrics.registerFontFamily(
            'Ping',
            normal='PingRegular',
            bold='PingBold',
            italic='PingRegular',
            boldItalic='PingBold'
        )
    except Exception:
        pass

    return True

def safe_set_font(c: canvas.Canvas, name: str, size: float):
    """
    Sets a ReportLab font safely, with fallback to Helvetica/Helvetica-Bold.
    """
    try:
        c.setFont(name, size)
    except Exception:
        fallback = "Helvetica-Bold" if "Bold" in name else "Helvetica"
        c.setFont(fallback, size)

# Call registration once at app start
register_fonts_once()

# -----------------------
# Helpers
# -----------------------
def now_local_str() -> str:
    return datetime.now(LOCAL_TZ).strftime("%Y-%m-%d %H:%M")

def today_local_date() -> str:
    return datetime.now(LOCAL_TZ).strftime("%Y%m%d")

def _join_root(root: str, path: str) -> str:
    """
    Mirrors the helper name from your stack trace.
    """
    if not root:
        return path
    return f"{root.rstrip('/')}/{path.lstrip('/')}"

# -----------------------
# PDF Generators
# -----------------------
# These sizes are examples; change to match your brand layout
EN_NAME_SIZE = 18
EN_ROLE_SIZE = 11
EN_META_SIZE = 9
MARGIN = 20 * mm

def signature_en_pdf(person: dict) -> bytes:
    """
    Builds a simple English signature PDF and returns raw bytes.
    Uses safe_set_font() to avoid crashes when custom font missing.
    """
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4

    # Example content layout
    name = f"{person.get('FirstName','').strip()} {person.get('LastName','').strip()}".strip()
    role = person.get('Title') or person.get('Position') or ""
    dept = person.get('Department') or ""
    phone = person.get('Phone') or ""
    email = person.get('Email') or ""
    company = person.get('Company') or "Alraedah Finance"

    y = height - MARGIN

    # Company (top-left)
    safe_set_font(c, "PingBold", 12)
    c.drawString(MARGIN, y, company)
    y -= 15

    # Name
    safe_set_font(c, "PingBold", EN_NAME_SIZE)
    c.drawString(MARGIN, y, name or "Employee Name")
    y -= 18

    # Role / Dept
    safe_set_font(c, "PingRegular", EN_ROLE_SIZE)
    c.drawString(MARGIN, y, role)
    y -= 14
    if dept:
        c.drawString(MARGIN, y, dept)
        y -= 14

    # Meta (phone / email)
    safe_set_font(c, "PingRegular", EN_META_SIZE)
    if phone:
        c.drawString(MARGIN, y, f"Phone: {phone}")
        y -= 12
    if email:
        c.drawString(MARGIN, y, f"Email: {email}")
        y -= 12

    # Footer timestamp
    safe_set_font(c, "PingRegular", 8)
    c.drawString(MARGIN, 10 * mm, f"Generated: {now_local_str()} (Asia/Riyadh)")

    c.showPage()
    c.save()
    buf.seek(0)
    return buf.getvalue()

# (Optional) Arabic signature builder — stub to extend later
def signature_ar_pdf(person: dict) -> bytes:
    """
    Placeholder for an Arabic signature PDF if needed later.
    Currently returns the EN PDF; replace with Arabic layout + shaping.
    """
    return signature_en_pdf(person)

# -----------------------
# ZIP Packager (mirrors your function name)
# -----------------------
def write_full_package_to_zip(zipf: zipfile.ZipFile, person: dict, root: str = ""):
    """
    Writes the full package for one person into an open ZipFile.
    Mirrors the function signature from your stack trace:
    write_full_package_to_zip(zipf, person)

    - Signature_EN.pdf
    - Signature_AR.pdf (placeholder)
    - (Add other artifacts if needed)
    """
    first = (person.get("FirstName") or "").strip()
    last = (person.get("LastName") or "").strip()
    folder = (f"{first}_{last}".strip() or "Employee").replace(" ", "_")

    # English signature
    en_pdf = signature_en_pdf(person)
    zipf.writestr(_join_root(root, f"{folder}/Signature_EN.pdf"), en_pdf)

    # Arabic signature (placeholder)
    ar_pdf = signature_ar_pdf(person)
    zipf.writestr(_join_root(root, f"{folder}/Signature_AR.pdf"), ar_pdf)

    # Example: add a tiny README per person
    readme = f"Package for {first} {last}\nGenerated: {now_local_str()} (Asia/Riyadh)\n"
    zipf.writestr(_join_root(root, f"{folder}/README.txt"), readme)

# -----------------------
# Streamlit UI
# -----------------------
st.markdown("Upload an Excel file and generate a ZIP with per-employee Signature PDFs.")

with st.expander("Excel Format (expected columns)"):
    st.write(
        "- FirstName (required)\n"
        "- LastName (required)\n"
        "- Title / Position\n"
        "- Department\n"
        "- Phone\n"
        "- Email\n"
        "- Company (optional; defaults to Alraedah Finance)\n"
    )

uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
package_name = st.text_input("Package name (root folder in ZIP)", value=f"Signatures_{today_local_date()}")

if st.button("Generate ZIP"):
    if not uploaded:
        st.error("Please upload an Excel file.")
    else:
        try:
            df = pd.read_excel(uploaded)
        except Exception as e:
            st.error(f"Could not read Excel: {e}")
            st.stop()

        required = ["FirstName", "LastName"]
        missing_req = [c for c in required if c not in df.columns]
        if missing_req:
            st.error(f"Missing required columns: {', '.join(missing_req)}")
            st.stop()

        # Build ZIP in memory
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            count = 0
            for _, row in df.iterrows():
                person = {k: ("" if pd.isna(v) else str(v)) for k, v in row.to_dict().items()}
                write_full_package_to_zip(zf, person, root=package_name)
                count += 1

            # Summary
            summary = (
                "Signature Package Summary\n"
                f"Generated: {now_local_str()} (Asia/Riyadh)\n"
                f"Root Folder: {package_name}\n"
                f"Total Employees: {count}\n"
            )
            zf.writestr(f"{package_name}/SUMMARY.txt", summary)

        zip_buf.seek(0)
        st.success("ZIP generated successfully.")
        st.download_button(
            "Download ZIP",
            data=zip_buf,
            file_name=f"{package_name}.zip",
            mime="application/zip"
        )

st.markdown(
    """
    <hr style="margin-top:40px;">
    <div style="text-align:center; color:gray; font-size:13px;">
      © 2025 Alraedah — Font-Safe Demo. If your custom fonts are missing, PDFs fall back to Helvetica.
    </div>
    """,
    unsafe_allow_html=True
)
