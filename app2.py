# ============================================================
# Alraedah — Email Signature & Business Card Suite (Merged)
# ============================================================
# Features:
# - Two main tabs: Email Signature / Business Card
# - Each has Single + Batch (Excel) sub-tabs
# - Excel template: QR template + Arabic_Name + Role (with dropdown validation)
# - Role EN->AR mapping applied to business card Arabic area
# - Preview Signature (PNG) + simple Card mock (PNG) + full Card PDF (2 pages)
# - QR code generation (PNG/SVG) + VCF
# - Toggle: include QR on designs (custom color + size placeholder)
# - Multi-select download: 4 scenarios (Full Package / Cards Only / Signatures Only / Master Folder)
# - Per-employee folder naming: FirstName_LastName_Role
#
# Notes:
# - Design functions are placeholders (clearly marked) so you can plug your Figma specs later.
# - ReportLab is used for the Business Card PDF (front/back).
# - Keep dependencies in requirements.txt: streamlit, pandas, pillow, qrcode, reportlab, xlsxwriter, python-barcode[images]
# ============================================================

import io
import os
import zipfile
from datetime import datetime
from dataclasses import dataclass, asdict

import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import qrcode
import qrcode.image.svg as qrcode_svg
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

# Optional: if you want EAN/Code128 later
# from barcode import Code128
# from barcode.writer import ImageWriter

# -----------------------------
# Page Config & Constants
# -----------------------------
st.set_page_config(
    page_title="Alraedah — Signature & Business Card Suite",
    page_icon="🧩",
    layout="wide",
)

TODAY = datetime.now().strftime("%Y-%m-%d")
APP_NAME = "Alraedah Suite"

# Example roles (English) — extend as needed
ROLES_EN = [
    "Chief Executive Officer",
    "Chief Marketing Officer",
    "Marketing Manager",
    "Relationship Manager",
    "Sales Specialist",
    "Customer Success Manager",
    "Head of HR",
    "HR Specialist",
    "IT Manager",
    "Software Engineer",
]

# English -> Arabic mapping for card Arabic line
ROLES_AR_MAP = {
    "Chief Executive Officer": "الرئيس التنفيذي",
    "Chief Marketing Officer": "الرئيس التنفيذي للتسويق",
    "Marketing Manager": "مدير التسويق",
    "Relationship Manager": "مدير علاقات",
    "Sales Specialist": "أخصائي مبيعات",
    "Customer Success Manager": "مدير نجاح العملاء",
    "Head of HR": "رئيس الموارد البشرية",
    "HR Specialist": "أخصائي موارد بشرية",
    "IT Manager": "مدير تقنية المعلومات",
    "Software Engineer": "مهندس برمجيات",
}

# --------------------------------
# Data structure for a single user
# --------------------------------
@dataclass
class Person:
    FirstName: str
    LastName: str
    Arabic_Name: str
    Role: str
    Email: str
    Phone: str
    Mobile: str
    Company: str
    Department: str
    Website: str
    LinkedIn: str
    Address: str

    @property
    def full_name_en(self):
        return f"{self.FirstName} {self.LastName}".strip()

    @property
    def safe_slug(self):
        base = f"{self.FirstName}_{self.LastName}_{self.Role}".replace(" ", "_")
        return "".join(ch for ch in base if ch.isalnum() or ch in ("_", "-"))

# -----------------------------
# Helpers
# -----------------------------
def vcard_from_person(p: Person) -> str:
    # Minimal vCard 3.0 — extend if needed
    lines = [
        "BEGIN:VCARD",
        "VERSION:3.0",
        f"N:{p.LastName};{p.FirstName};;;",
        f"FN:{p.full_name_en}",
        f"TITLE:{p.Role}",
        f"ORG:{p.Company}",
    ]
    if p.Email:
        lines.append(f"EMAIL;TYPE=INTERNET:{p.Email}")
    if p.Mobile:
        lines.append(f"TEL;TYPE=CELL:{p.Mobile}")
    if p.Phone:
        lines.append(f"TEL;TYPE=WORK,VOICE:{p.Phone}")
    if p.Website:
        lines.append(f"URL:{p.Website}")
    if p.Address:
        lines.append(f"ADR;TYPE=WORK:;;{p.Address};;;;")
    lines.append("END:VCARD")
    return "\n".join(lines)

def make_qr_png_bytes(data: str, fill_color="black", back_color="white") -> bytes:
    qr = qrcode.QRCode(
        version=4,  # adjust automatically if needed (set None and add fit=True)
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=2,
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color=fill_color, back_color=back_color).convert("RGB")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

def make_qr_svg_bytes(data: str) -> bytes:
    img = qrcode.make(data, image_factory=qrcode_svg.SvgImage)
    buf = io.BytesIO()
    img.save(buf)
    return buf.getvalue()

def make_signature_png(p: Person, include_qr: bool, qr_color: str) -> bytes:
    """
    Placeholder Signature Renderer (PNG)
    - Replace with your finalized design later.
    - Bilingual layout idea: EN left / AR right, dark/light styling to be added.
    """
    W, H = 1200, 480  # high-res base
    img = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)

    # Fonts: use system defaults; swap with brand fonts in production
    try:
        font_title = ImageFont.truetype("Arial.ttf", 42)
        font_text = ImageFont.truetype("Arial.ttf", 28)
        font_ar = ImageFont.truetype("Arial.ttf", 36)
    except:
        font_title = ImageFont.load_default()
        font_text = ImageFont.load_default()
        font_ar = ImageFont.load_default()

    # Left (EN)
    xL, y = 40, 40
    draw.text((xL, y), p.full_name_en, fill="black", font=font_title); y += 60
    draw.text((xL, y), p.Role, fill="gray", font=font_text); y += 44
    draw.text((xL, y), p.Company, fill="black", font=font_text); y += 44
    if p.Email:   draw.text((xL, y), f"Email: {p.Email}", fill="black", font=font_text); y += 36
    if p.Mobile:  draw.text((xL, y), f"Mobile: {p.Mobile}", fill="black", font=font_text); y += 36
    if p.Phone:   draw.text((xL, y), f"Phone: {p.Phone}", fill="black", font=font_text); y += 36
    if p.Website: draw.text((xL, y), f"Web: {p.Website}", fill="black", font=font_text); y += 36
    if p.LinkedIn:draw.text((xL, y), f"LinkedIn: {p.LinkedIn}", fill="black", font=font_text); y += 36

    # Right (AR)
    # Minimal Arabic block (reversed area). Substitute right-to-left properly later with bidi/arabic-reshaper if needed.
    xR = W - 560
    yR = 60
    role_ar = ROLES_AR_MAP.get(p.Role, p.Role)
    draw.rectangle([xR - 20, 30, W - 30, H - 30], outline="#1D4283", width=3)
    draw.text((xR, yR), p.Arabic_Name or "", fill="black", font=font_ar); yR += 60
    draw.text((xR, yR), role_ar, fill="#1D4283", font=font_ar); yR += 60
    draw.text((xR, yR), "الرائدة للتمويل", fill="black", font=font_ar)

    # Optional QR (colored)
    if include_qr:
        qr_png = Image.open(io.BytesIO(make_qr_png_bytes(vcard_from_person(p), fill_color=qr_color)))
        qr_size = 180  # placeholder; wire to UI later
        qr_png = qr_png.resize((qr_size, qr_size), Image.LANCZOS)
        img.paste(qr_png, (W - qr_size - 40, H - qr_size - 40))

    out = io.BytesIO()
    img.save(out, format="PNG")
    return out.getvalue()

def make_business_card_pdf(p: Person, include_qr: bool, qr_color: str,
                           bg_front_path="card_front.png", bg_back_path="card_back.png") -> bytes:
    """
    Business Card PDF (2 pages: front/back) using PNG backgrounds.
    - bg_front_path and bg_back_path: file paths to your PNG background designs.
    - Overlay text + optional QR code on them.
    """
    CARD_W, CARD_H = 90 * mm, 54 * mm
    margin = 6 * mm

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(CARD_W, CARD_H))

    # ---------- FRONT ----------
    if os.path.exists(bg_front_path):
        bg_img = ImageReader(bg_front_path)
        c.drawImage(bg_img, 0, 0, CARD_W, CARD_H)

    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(margin, CARD_H - margin - 12, p.full_name_en)
    c.setFont("Helvetica", 9)
    c.drawString(margin, CARD_H - margin - 26, p.Role)
    c.drawString(margin, CARD_H - margin - 40, p.Company)
    y = CARD_H - margin - 58
    if p.Email:   c.drawString(margin, y, f"Email: {p.Email}"); y -= 12
    if p.Mobile:  c.drawString(margin, y, f"Mobile: {p.Mobile}"); y -= 12
    if p.Website: c.drawString(margin, y, f"Web: {p.Website}"); y -= 12

    if include_qr:
        qr_png = make_qr_png_bytes(vcard_from_person(p), fill_color=qr_color)
        qr_img = Image.open(io.BytesIO(qr_png))
        qr_reader = ImageReader(qr_img)
        qr_w = 18 * mm  # final printed size
        c.drawImage(qr_reader, CARD_W - margin - qr_w, margin, qr_w, qr_w, mask='auto')

    c.showPage()

    # ---------- BACK ----------
    if os.path.exists(bg_back_path):
        bg_img = ImageReader(bg_back_path)
        c.drawImage(bg_img, 0, 0, CARD_W, CARD_H)

    role_ar = ROLES_AR_MAP.get(p.Role, p.Role)
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 12)
    c.drawRightString(CARD_W - margin, CARD_H - margin - 16, p.Arabic_Name or "")
    c.setFont("Helvetica", 11)
    c.drawRightString(CARD_W - margin, CARD_H - margin - 30, role_ar)
    c.drawRightString(CARD_W - margin, CARD_H - margin - 44, "الرائدة للتمويل")

    c.showPage()
    c.save()
    return buf.getvalue()


def people_from_excel(file) -> list:
    df = pd.read_excel(file).fillna("")
    # Expected columns (from template): keep in sync
    expected = [
        "FirstName","LastName","Arabic_Name","Role","Email",
        "Phone","Mobile","Company","Department","Website","LinkedIn","Address"
    ]
    missing = [c for c in expected if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in Excel: {missing}")

    people = []
    for _, row in df.iterrows():
        people.append(Person(
            FirstName=str(row["FirstName"]).strip(),
            LastName=str(row["LastName"]).strip(),
            Arabic_Name=str(row["Arabic_Name"]).strip(),
            Role=str(row["Role"]).strip(),
            Email=str(row["Email"]).strip(),
            Phone=str(row["Phone"]).strip(),
            Mobile=str(row["Mobile"]).strip(),
            Company=str(row["Company"]).strip(),
            Department=str(row["Department"]).strip(),
            Website=str(row["Website"]).strip(),
            LinkedIn=str(row["LinkedIn"]).strip(),
            Address=str(row["Address"]).strip(),
        ))
    return people, df

def make_excel_template_bytes(roles_list: list) -> bytes:
    """
    Builds an Excel template with data validation (dropdown) for Role.
    """
    import xlsxwriter
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet("Data")
    meta = wb.add_worksheet("Roles")  # hidden list

    headers = [
        "FirstName","LastName","Arabic_Name","Role","Email",
        "Phone","Mobile","Company","Department","Website","LinkedIn","Address"
    ]
    for col, h in enumerate(headers):
        ws.write(0, col, h)

    # Write roles in hidden sheet
    for idx, role in enumerate(roles_list, start=0):
        meta.write(idx, 0, role)

    # Define a named range for validation
    wb.define_name('RolesList', f'=Roles!$A$1:$A${len(roles_list)}')

    # Apply data validation on Role column (index 3)
    ws.data_validation(1, 3, 1000, 3, {
        'validate': 'list',
        'source': '=RolesList'
    })

    # Hide Roles sheet
    meta.hide()

    wb.close()
    output.seek(0)
    return output.getvalue()

# -----------------------------
# Packaging / ZIP Scenarios
# -----------------------------
SCENARIOS = [
    "Full Package (per-employee folders)",
    "Business Cards Only (flat folder)",
    "Email Signatures Only (flat folder)",
    "Master Folder (subfolders of 1–3)",
]

def build_zip_for_people(people: list, include_qr_on_design: bool, qr_color: str, chosen_scenarios: list) -> bytes:
    """
    Builds a ZIP according to selected scenarios.
    Folder naming: FirstName_LastName_Role
    """
    zbuf = io.BytesIO()
    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
        base_full = f"Generation_{TODAY}/"
        base_cards = f"Business_Cards_{TODAY}/"
        base_sigs = f"Email_Signatures_{TODAY}/"
        base_master = f"Master_{TODAY}/"

        # If scenario 4 (master), create sub-bases
        m_full = base_master + "Full_Package/"
        m_cards = base_master + "Business_Cards_Only/"
        m_sigs = base_master + "Email_Signatures_Only/"

        for p in people:
            slug = p.safe_slug

            # Build content once per person
            # QR assets
            vcf_text = vcard_from_person(p)
            qr_png = make_qr_png_bytes(vcf_text)  # default black for QR assets folder
            qr_svg = make_qr_svg_bytes(vcf_text)

            # Signature PNG (with optional QR overlay color)
            sig_png = make_signature_png(p, include_qr_on_design, qr_color)

            # Business Card PDF
            card_pdf = make_business_card_pdf(p, include_qr_on_design, qr_color)

            # Scenario 1 — Full Package (per-employee folders)
            if SCENARIOS[0] in chosen_scenarios:
                emp_dir = f"{base_full}{slug}/"
                qr_dir = emp_dir + "QR_Codes/"
                z.writestr(qr_dir, "")  # ensure folder
                z.writestr(emp_dir, "")
                z.writestr(qr_dir + f"{slug}.png", qr_png)
                z.writestr(qr_dir + f"{slug}.svg", qr_svg)
                z.writestr(qr_dir + f"{slug}.vcf", vcf_text)
                z.writestr(emp_dir + "Email_Signature.png", sig_png)
                z.writestr(emp_dir + "Business_Card.pdf", card_pdf)

            # Scenario 2 — Business Cards Only (flat folder)
            if SCENARIOS[1] in chosen_scenarios:
                z.writestr(base_cards + f"{slug}.pdf", card_pdf)

            # Scenario 3 — Email Signatures Only (flat folder)
            if SCENARIOS[2] in chosen_scenarios:
                z.writestr(base_sigs + f"{slug}.png", sig_png)

            # Scenario 4 — Master Folder with subfolders of 1–3
            if SCENARIOS[3] in chosen_scenarios:
                # Full Package inside Master
                emp_dir = f"{m_full}{slug}/"
                qr_dir = emp_dir + "QR_Codes/"
                z.writestr(qr_dir, ""); z.writestr(emp_dir, "")
                z.writestr(qr_dir + f"{slug}.png", qr_png)
                z.writestr(qr_dir + f"{slug}.svg", qr_svg)
                z.writestr(qr_dir + f"{slug}.vcf", vcf_text)
                z.writestr(emp_dir + "Email_Signature.png", sig_png)
                z.writestr(emp_dir + "Business_Card.pdf", card_pdf)
                # Cards-only & Sigs-only inside Master
                z.writestr(m_cards + f"{slug}.pdf", card_pdf)
                z.writestr(m_sigs + f"{slug}.png", sig_png)

    zbuf.seek(0)
    return zbuf.getvalue()

# -----------------------------
# UI — Forms
# -----------------------------
def single_person_form(defaults=None):
    col1, col2, col3 = st.columns(3)
    with col1:
        FirstName = st.text_input("First Name", (defaults or {}).get("FirstName", ""))
        Arabic_Name = st.text_input("Arabic Name", (defaults or {}).get("Arabic_Name", ""))
        Phone = st.text_input("Phone (work)", (defaults or {}).get("Phone", ""))
        Company = st.text_input("Company", (defaults or {}).get("Company", "Alraedah Finance"))
    with col2:
        LastName = st.text_input("Last Name", (defaults or {}).get("LastName", ""))
        Role = st.selectbox("Role (EN)", ROLES_EN, index=ROLES_EN.index((defaults or {}).get("Role", ROLES_EN[0])) if (defaults and (defaults.get("Role") in ROLES_EN)) else 0)
        Mobile = st.text_input("Mobile", (defaults or {}).get("Mobile", ""))
        Department = st.text_input("Department", (defaults or {}).get("Department", ""))
    with col3:
        Email = st.text_input("Email", (defaults or {}).get("Email", ""))
        Website = st.text_input("Website", (defaults or {}).get("Website", ""))
        LinkedIn = st.text_input("LinkedIn", (defaults or {}).get("LinkedIn", ""))
        Address = st.text_input("Address", (defaults or {}).get("Address", "Riyadh, Saudi Arabia"))

    p = Person(
        FirstName=FirstName, LastName=LastName, Arabic_Name=Arabic_Name, Role=Role,
        Email=Email, Phone=Phone, Mobile=Mobile, Company=Company, Department=Department,
        Website=Website, LinkedIn=LinkedIn, Address=Address
    )
    return p

def preview_signature_and_card(p: Person, include_qr: bool, qr_color: str):
    st.subheader("Preview")
    sig_png = make_signature_png(p, include_qr, qr_color)
    st.image(sig_png, caption="Email Signature Preview (PNG)", use_column_width=True)

    # For card, show a simple front mock as PNG (optional). We'll reuse signature image as proxy.
    st.info("Business Card preview uses a placeholder mock. Final design will match your brand layout.")
    st.image(sig_png, caption="Business Card Front (mock)", use_column_width=True)

# -----------------------------
# App Layout
# -----------------------------
st.title("🧩 Alraedah — Email Signature & Business Card Suite")

with st.expander("📥 Download Excel Template (with Role dropdown)"):
    tmp_bytes = make_excel_template_bytes(ROLES_EN)
    st.download_button("Download Excel Template (.xlsx)", data=tmp_bytes, file_name=f"alraedah_template_{TODAY}.xlsx")

# Global design toggles
st.sidebar.header("Design Options")
include_qr_on_design = st.sidebar.toggle("Include QR on designs", value=True)
qr_color = st.sidebar.color_picker("QR Color (for designs)", value="#000000")
download_choices = st.sidebar.multiselect("Download Scenarios (multi-select)", SCENARIOS, default=[SCENARIOS[0]])

tab_email, tab_card = st.tabs(["✉️ Email Signature Generator", "🪪 Business Card Generator"])

# ---------------- Email Signature Tab ----------------
with tab_email:
    sub_e_single, sub_e_batch = st.tabs(["Single", "Batch (Excel)"])

    with sub_e_single:
        st.markdown("Fill person info, preview, then generate.")
        person = single_person_form()
        if st.button("Preview Signature & Card (from this data)"):
            preview_signature_and_card(person, include_qr_on_design, qr_color)

        if st.button("Generate & Download (Single)"):
            data_zip = build_zip_for_people([person], include_qr_on_design, qr_color, download_choices)
            st.download_button(
                "Download ZIP",
                data=data_zip,
                file_name=f"{APP_NAME}_Single_{TODAY}.zip",
                mime="application/zip",
            )

    with sub_e_batch:
        st.markdown("Upload the Excel filled from the template. You can then preview and generate all.")
        up = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
        if up:
            try:
                people, df = people_from_excel(up)
                st.success(f"Loaded {len(people)} records.")
                st.dataframe(df)
                if st.button("Generate & Download (Batch via Email tab)"):
                    data_zip = build_zip_for_people(people, include_qr_on_design, qr_color, download_choices)
                    st.download_button(
                        "Download ZIP",
                        data=data_zip,
                        file_name=f"{APP_NAME}_EmailBatch_{TODAY}.zip",
                        mime="application/zip",
                    )
            except Exception as e:
                st.error(f"Error: {e}")

# ---------------- Business Card Tab ----------------
with tab_card:
    sub_c_single, sub_c_batch = st.tabs(["Single", "Batch (Excel)"])

    with sub_c_single:
        st.markdown("Enter details for a single card. (Same data model as signature.)")
        person_c = single_person_form()
        if st.button("Preview Card & Signature (from this data)"):
            preview_signature_and_card(person_c, include_qr_on_design, qr_color)

        if st.button("Generate & Download (Single via Card tab)"):
            data_zip = build_zip_for_people([person_c], include_qr_on_design, qr_color, download_choices)
            st.download_button(
                "Download ZIP",
                data=data_zip,
                file_name=f"{APP_NAME}_CardSingle_{TODAY}.zip",
                mime="application/zip",
            )

    with sub_c_batch:
        st.markdown("Upload the same Excel template — roles are in a dropdown.")
        up2 = st.file_uploader("Upload Excel (.xlsx) for Cards", type=["xlsx"], key="cards_excel")
        if up2:
            try:
                people2, df2 = people_from_excel(up2)
                st.success(f"Loaded {len(people2)} records.")
                st.dataframe(df2)
                if st.button("Generate & Download (Batch via Card tab)"):
                    data_zip = build_zip_for_people(people2, include_qr_on_design, qr_color, download_choices)
                    st.download_button(
                        "Download ZIP",
                        data=data_zip,
                        file_name=f"{APP_NAME}_CardBatch_{TODAY}.zip",
                        mime="application/zip",
                    )
            except Exception as e:
                st.error(f"Error: {e}")

# ---------------- Footer / Help ----------------
st.markdown("---")
st.caption("© Alraedah — This is a functional skeleton. Replace placeholder renderers with your final brand designs (fonts, sizes, colors, grids).")