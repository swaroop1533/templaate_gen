import os
import base64
import requests
import pandas as pd
import xml.etree.ElementTree as ET
import subprocess
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader

# ===================== CONFIG ======================

# Paths
EXCEL_PATH = "/home/swaroop/Desktop/templaate_gen/Attendance_Card_GBM_On_17th_Jan_26 2.xlsx"
TEMPLATE_SVG_PATH = "/home/swaroop/Desktop/templaate_gen/Latest 16012026 (1) 1.svg"
FONT_TTF_PATH = "/home/swaroop/Desktop/templaate_gen/ARIALNB.TTF"

SVG_DIR = "output_cards_svg"       # generated SVG cards
PNG_DIR = "output_cards_png"       # high-res PNGs
OUTPUT_PDF = "final_generation.pdf"

INKSCAPE_CMD = "inkscape"          # adjust if needed

# Excel columns
COL_QR = "GENERATED_QRCODE"
COL_NAME = "NAME"
COL_ID = "MEMBER_ID"
COL_MOBILE = "MOBILE_NO"

# Layout inside the card (percentages of SVG width/height)
QR_BOX_PERC = (0.47, 0.10, 0.97, 0.48)   # QR area

TEXT_CENTER_X_PERC = 0.50                # text center X
TEXT_Y_START_PERC = 0.545                # first line Y
LINE_SPACING_PERC = 0.06                 # vertical gap between lines

NAME_FONT_PERC = 0.04
OTHER_FONT_PERC = 0.04

FONT_FAMILY = "ArialNarrowBold"
TEXT_COLOR = "#3a53a4"

# PDF grid layout
CARDS_PER_ROW = 3
ROWS_PER_PAGE = 3
CARDS_PER_PAGE = CARDS_PER_ROW * ROWS_PER_PAGE

PNG_DPI = 900    # export DPI for PNG

# Card box size in PDF (points) – match your SVG size (~170 x 227)
CARD_W_PT = 170.28   # width of card box
CARD_H_PT = 226.97   # height of card box

# ====================================================
# =============== SVG GENERATION PART ================
# ====================================================

def get_card_size(root):
    """Get card width/height from viewBox='0 0 w h'."""
    view_box = root.attrib.get("viewBox")
    if not view_box:
        raise ValueError("SVG template must have a viewBox attribute.")

    parts = view_box.strip().split()
    if len(parts) != 4:
        raise ValueError(f"Unexpected viewBox format: {view_box}")

    width = float(parts[2])
    height = float(parts[3])
    return width, height


def get_ns_helpers(root):
    """Return namespace helpers and register namespaces."""
    if root.tag.startswith("{"):
        svg_ns = root.tag[1:].split("}")[0]
    else:
        svg_ns = "http://www.w3.org/2000/svg"

    xlink_ns = "http://www.w3.org/1999/xlink"

    ET.register_namespace("", svg_ns)
    ET.register_namespace("xlink", xlink_ns)

    def t(name: str) -> str:
        return f"{{{svg_ns}}}{name}"

    return svg_ns, xlink_ns, t


def embed_font_face(root, font_path, font_family_name):
    """Embed a TTF into the SVG using @font-face so viewer uses it."""
    svg_ns, xlink_ns, t = get_ns_helpers(root)  # noqa: F841

    with open(font_path, "rb") as f:
        font_data = f.read()
    b64_font = base64.b64encode(font_data).decode("ascii")

    defs = root.find(f".//{{{svg_ns}}}defs")
    if defs is None:
        defs = ET.SubElement(root, t("defs"))

    style_el = defs.find(f".//{{{svg_ns}}}style")
    if style_el is None:
        style_el = ET.SubElement(defs, t("style"), {"type": "text/css"})
        style_el.text = ""

    css = style_el.text or ""
    css += f"""
@font-face {{
    font-family: '{font_family_name}';
    src: url("data:font/truetype;base64,{b64_font}") format("truetype");
    font-weight: bold;
}}
"""
    style_el.text = css


def add_text(root, tag, text, x, y, font_size, bold=False, center=False):
    attrs = {
        "x": str(x),
        "y": str(y),
        "font-size": str(font_size),
        "font-family": FONT_FAMILY,
        "fill": TEXT_COLOR,
    }
    if bold:
        attrs["font-weight"] = "bold"
    if center:
        attrs["text-anchor"] = "middle"

    el = ET.SubElement(root, tag("text"), attrs)
    el.text = text
    return el


def add_qr_png_image(root, tag, xlink_ns, png_url, x, y, w, h):
    """Download PNG QR and embed into SVG as <image> with data URI."""
    resp = requests.get(png_url, timeout=20)
    resp.raise_for_status()

    b64_data = base64.b64encode(resp.content).decode("ascii")
    href_val = f"data:image/png;base64,{b64_data}"

    attrs = {
        "x": str(x),
        "y": str(y),
        "width": str(w),
        "height": str(h),
        f"{{{xlink_ns}}}href": href_val,
    }
    ET.SubElement(root, tag("image"), attrs)


def generate_svg_cards():
    """Read Excel and generate per-member SVG cards in SVG_DIR."""
    os.makedirs(SVG_DIR, exist_ok=True)

    df = pd.read_excel(EXCEL_PATH)

    for col in [COL_QR, COL_NAME, COL_ID, COL_MOBILE]:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found in Excel.")

    template_tree = ET.parse(TEMPLATE_SVG_PATH)
    template_root = template_tree.getroot()
    card_w, card_h = get_card_size(template_root)
    print(f"Card size from SVG viewBox: {card_w} x {card_h}")

    for idx, row in df.iterrows():
        raw_qr = row.get(COL_QR, None)
        if raw_qr is None or pd.isna(raw_qr):
            print(f"Row {idx}: QR column empty, skipping.")
            continue

        qr_url = str(raw_qr).strip()
        name = str(row[COL_NAME]).strip()
        member_id = str(row[COL_ID]).strip()
        mobile = str(row[COL_MOBILE]).strip()

        print(f"Row {idx}: generating for {name} ({member_id}) → {qr_url}")

        # fresh copy of template
        root = ET.fromstring(ET.tostring(template_root))
        _, xlink_ns, t = get_ns_helpers(root)

        # embed font
        embed_font_face(root, FONT_TTF_PATH, FONT_FAMILY)

        # layout calculations
        text_center_x = card_w * TEXT_CENTER_X_PERC
        text_y_start = card_h * TEXT_Y_START_PERC
        line_spacing = card_h * LINE_SPACING_PERC

        name_font_size = card_h * NAME_FONT_PERC
        other_font_size = card_h * OTHER_FONT_PERC

        qr_x1 = card_w * QR_BOX_PERC[0]
        qr_y1 = card_h * QR_BOX_PERC[1]
        qr_x2 = card_w * QR_BOX_PERC[2]
        qr_y2 = card_h * QR_BOX_PERC[3]
        qr_w = qr_x2 - qr_x1
        qr_h = qr_y2 - qr_y1

        # add QR
        add_qr_png_image(
            root, t, xlink_ns,
            png_url=qr_url,
            x=qr_x1, y=qr_y1,
            w=qr_w, h=qr_h,
        )

        # text lines
        line1 = f"NAME : {name}"
        line2 = f"MEMBER ID No : {member_id}"
        line3 = f"MOB : {mobile}"

        y = text_y_start
        add_text(root, t, line1, text_center_x, y, name_font_size, bold=True, center=True)
        y += line_spacing
        add_text(root, t, line2, text_center_x, y, other_font_size, bold=True, center=True)
        y += line_spacing
        add_text(root, t, line3, text_center_x, y, other_font_size, bold=True, center=True)

        # save SVG
        safe_name = "".join(
            c for c in name.replace(" ", "_") if c.isalnum() or c in ["_", "-"]
        )
        filename = f"{member_id}_{safe_name}.svg"
        out_path = os.path.join(SVG_DIR, filename)

        ET.ElementTree(root).write(out_path, encoding="utf-8", xml_declaration=True)
        print(f"Saved SVG: {out_path}")

    print("SVG generation done.")


# ====================================================
# =============== PNG + PDF PART =====================
# ====================================================

def ensure_png_dir():
    os.makedirs(PNG_DIR, exist_ok=True)


def convert_svg_to_png():
    """Convert SVG cards to high-res PNGs using Inkscape."""
    svg_files = sorted(f for f in os.listdir(SVG_DIR) if f.lower().endswith(".svg"))
    if not svg_files:
        print("No SVGs found in", SVG_DIR)
        return []

    png_paths = []
    for svg_name in svg_files:
        svg_path = os.path.join(SVG_DIR, svg_name)
        base = os.path.splitext(svg_name)[0]
        png_path = os.path.join(PNG_DIR, base + ".png")

        print(f"Converting SVG → PNG: {svg_path} → {png_path}")
        cmd = [
            INKSCAPE_CMD,
            svg_path,
            "--export-type=png",
            f"--export-filename={png_path}",
            f"--export-dpi={PNG_DPI}",
        ]
        subprocess.run(cmd, check=True)
        png_paths.append(png_path)

    return png_paths


def build_pdf(png_paths):
    """Place each card into a fixed 170x227 pt box, 3x3 per A4 page."""
    if not png_paths:
        print("No PNGs found to add to PDF.")
        return

    c = canvas.Canvas(OUTPUT_PDF, pagesize=A4)
    page_w, page_h = A4  # in points

    total_cards_w = CARDS_PER_ROW * CARD_W_PT
    total_cards_h = ROWS_PER_PAGE * CARD_H_PT

    gap_x = (page_w - total_cards_w) / (CARDS_PER_ROW + 1)
    gap_y = (page_h - total_cards_h) / (ROWS_PER_PAGE + 1)

    for idx, png_file in enumerate(png_paths):
        pos = idx % CARDS_PER_PAGE

        if pos == 0 and idx != 0:
            c.showPage()

        col = pos % CARDS_PER_ROW
        row = pos // CARDS_PER_ROW

        row_from_bottom = ROWS_PER_PAGE - 1 - row

        # bottom-left of card box
        box_x = gap_x + col * (CARD_W_PT + gap_x)
        box_y = gap_y + row_from_bottom * (CARD_H_PT + gap_y)

        img = ImageReader(png_file)
        img_w_px, img_h_px = img.getSize()

        # scale image to fit box
        scale = min(CARD_W_PT / img_w_px, CARD_H_PT / img_h_px)
        draw_w = img_w_px * scale
        draw_h = img_h_px * scale

        draw_x = box_x + (CARD_W_PT - draw_w) / 2
        draw_y = box_y + (CARD_H_PT - draw_h) / 2

        print(f"Placing {png_file} at box ({box_x:.1f}, {box_y:.1f}) size {CARD_W_PT}x{CARD_H_PT}")
        c.drawImage(img, draw_x, draw_y, draw_w, draw_h)

    c.save()
    print("PDF saved:", OUTPUT_PDF)


# ====================================================
# ==================== MAIN ==========================
# ====================================================

def main():
    # 1) Generate SVG cards from Excel
    generate_svg_cards()

    # 2) Convert SVG → high-res PNG
    ensure_png_dir()
    png_paths = convert_svg_to_png()

    # 3) Build 3x3 PDF
    build_pdf(png_paths)


if __name__ == "__main__":
    main()
