# app.py

import io
import os
import re
import tempfile

import pandas as pd
import requests
import streamlit as st
from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import inch
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import (
    Image,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Logo and custom font URLs
LOGO_URL = (
    "https://www.carnegiehighered.com/wp-content/uploads/2021/11/"
    "Twitter-Image-2-2021.png"
)
FONTS = {
    "Barlow":  "https://raw.githubusercontent.com/scooter7/ExcelBudgetBox/main/fonts/Barlow-Regular.ttf",
    "DMSerif": "https://raw.githubusercontent.com/scooter7/ExcelBudgetBox/main/fonts/DMSerifDisplay-Regular.ttf",
}

# Register custom fonts
for name, url in FONTS.items():
    r = requests.get(url)
    path = tempfile.NamedTemporaryFile(delete=False, suffix=".ttf").name
    with open(path, "wb") as f:
        f.write(r.content)
    pdfmetrics.registerFont(TTFont(name, path))

st.set_page_config(page_title="Proposal → PDF", layout="wide")


def load_and_prepare_dataframe(uploaded_file):
    """
    Load Excel/CSV using the 2nd row as header,
    then rename column A to 'Service'.
    """
    fn = uploaded_file.name.lower()
    if fn.endswith((".xls", ".xlsx")):
        df = pd.read_excel(uploaded_file, header=1)
    else:
        df = pd.read_csv(uploaded_file, header=1)
    first = df.columns[0]
    if first != "Service":
        df = df.rename(columns={first: "Service"})
    return df


def transform_service_column(df):
    """
    Clean the 'Service' column: remove leading 'XX:',
    drop slash+suffix, and strip out any parentheses.
    """
    def clean(s):
        s = str(s)
        s = re.sub(r'^..:', '', s)       # drop leading XX:
        s = s.split('/', 1)[0]           # drop slash and after
        s = re.sub(r'\(.*?\)', '', s)    # remove parentheses content
        return s.strip()
    df["Service"] = df["Service"].fillna("").apply(clean)
    return df


def replace_est(df):
    """
    Rename Est. → Estimated in column headers and cell values.
    """
    df.columns = [c.replace("Est.", "Estimated").strip() for c in df.columns]
    return df.applymap(lambda v: (v.replace("Est.", "Estimated") if isinstance(v, str) else v))


def split_tables(df):
    """
    Split the DataFrame into segments by legacy 'Total' rows.
    Each segment is named by its first non-header Service value.
    Returns a list of dicts: {'name': str, 'df': DataFrame}.
    """
    segments = []
    start = 0
    for i, svc in enumerate(df["Service"]):
        if str(svc).strip().lower() == "total":
            seg = df.iloc[start : i + 1].reset_index(drop=True)
            # derive name from first non-header Service
            name = next(
                (x for x in seg["Service"] if x and x.strip().lower() != "service"), 
                f"Table {len(segments)+1}"
            )
            segments.append({"name": name, "df": seg})
            start = i + 1
    # trailing rows without a Total
    if start < len(df):
        seg = df.iloc[start:].reset_index(drop=True)
        name = next(
            (x for x in seg["Service"] if x and x.strip().lower() != "service"),
            f"Table {len(segments)+1}"
        )
        segments.append({"name": name, "df": seg})
    return segments


def calculate_and_insert_totals(seg_df):
    """
    Pull the Excel's original Total row values, drop it,
    then re-append it so the PDF shows the real totals.
    """
    df = seg_df.copy()
    # 1) locate original Total row
    mask = df["Service"].str.strip().str.lower() == "total"
    orig = df.loc[mask].iloc[0] if mask.any() else None

    # 2) drop all legacy footers
    drop_keys = {
        "total",
        "est. conversions", "estimated conversions",
        "est. impressions", "estimated impressions",
    }
    df = df.loc[~df["Service"].str.strip().str.lower().isin(drop_keys)].reset_index(drop=True)

    # 3) build new Total row from orig values
    total = {c: "" for c in df.columns}
    total["Service"] = "Total"
    if orig is not None:
        for c in df.columns:
            val = orig.get(c)
            if pd.notna(val):
                total[c] = val
    return pd.concat([df, pd.DataFrame([total])], ignore_index=True)


def make_pdf(segments, title):
    """
    Render each named DataFrame segment into a single 11x17 PDF.
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=(17 * inch, 11 * inch),
        leftMargin=0.5 * inch,
        rightMargin=0.5 * inch,
        topMargin=0.5 * inch,
        bottomMargin=0.5 * inch,
    )

    styles = getSampleStyleSheet()
    hdr_style = ParagraphStyle(
        "hdr", parent=styles["BodyText"], fontName="DMSerif", fontSize=8, leading=9, alignment=1
    )
    body_style = ParagraphStyle(
        "bod", parent=styles["BodyText"], fontName="Barlow", fontSize=7, leading=8, alignment=0
    )

    elems = []
    # Add logo ×3 size
    resp = requests.get(LOGO_URL)
    logo_img = PILImage.open(io.BytesIO(resp.content))
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    logo_img.save(tmp.name)
    elems.append(Image(tmp.name, width=4.5 * inch, height=1.5 * inch))
    elems.append(Spacer(1, 12))

    # Title
    elems.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
    elems.append(Spacer(1, 12))

    # Render each segment
    for seg in segments:
        name = seg["name"]
        df = seg["df"]

        # Drop blank and repeated header rows
        df = df[df["Service"].notna() & df["Service"].str.strip().astype(bool)]
        df = df.loc[
            ~(
                (df["Service"].str.strip().str.lower() == "service")
                & (df.get("Description", "").str.strip().str.lower() == "description")
            )
        ].reset_index(drop=True)

        # Re-append true totals
        df = calculate_and_insert_totals(df)

        # Format dates and currency
        for col in df.columns:
            if "date" in col.lower():
                df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%m/%d/%Y")
        for col in ("Monthly Amount", "Item Total"):
            if col in df.columns:
                df[col] = (
                    pd.to_numeric(df[col], errors="coerce")
                    .fillna(0)
                    .map(lambda x: f"${x:,.0f}")
                )

        # Segment heading
        elems.append(Paragraph(f"<b>{name}</b>", styles["Heading2"]))
        elems.append(Spacer(1, 6))

        # Build table data
        header_cells = [Paragraph(c, hdr_style) for c in df.columns]
        data = [header_cells]
        for row in df.itertuples(index=False):
            cells = []
            for c, v in zip(df.columns, row):
                txt = "" if v is None else str(v)
                if "<link" in txt or c in ("Service", "Description", "Notes"):
                    cells.append(Paragraph(txt, body_style))
                else:
                    cells.append(txt)
            data.append(cells)

        col_widths = [doc.width * w for w in [0.12, 0.30, 0.06, 0.08, 0.08, 0.12, 0.12, 0.06, 0.06]]
        table = Table(data, colWidths=col_widths, repeatRows=1)
        table.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.4, colors.black),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                    ("FONTNAME", (0, 0), (-1, 0), "DMSerif"),
                    ("FONTNAME", (0, 1), (-1, -1), "Barlow"),
                    ("FONTSIZE", (0, 0), (-1, -1), 7),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("BACKGROUND", (0, len(data) - 1), (-1, len(data) - 1), colors.lightgrey),
                    ("FONTNAME", (0, len(data) - 1), (-1, len(data) - 1), "DMSerif"),
                ]
            )
        )
        elems.append(table)
        elems.append(Spacer(1, 24))

    doc.build(elems)
    buf.seek(0)
    return buf


def main():
    st.title("🔄 Proposal → PDF")
    uploaded = st.file_uploader("Upload Excel/CSV", type=["xls", "xlsx", "csv"])
    if not uploaded:
        return

    # Load & clean
    df = load_and_prepare_dataframe(uploaded)
    df = transform_service_column(df)
    df = replace_est(df)

    # Split into named segments
    segments = split_tables(df)

    # 1) Row removal per table
    table_names = [seg["name"] for seg in segments]
    to_remove = st.multiselect("Remove rows from which tables?", options=table_names)
    for seg in segments:
        if seg["name"] in to_remove:
            preview = seg["df"].reset_index().rename(columns={"index": "ID"})
            st.write(f"Rows in **{seg['name']}**:")
            st.dataframe(preview, use_container_width=True)
            ids = st.multiselect(
                f"Select IDs to drop from {seg['name']} (0-based):",
                options=preview["ID"].tolist(),
                key=f"drop_{seg['name']}",
            )
            if ids:
                seg["df"] = seg["df"].drop(index=ids).reset_index(drop=True)

    # 2) Column removal per table
    to_modify = st.multiselect("Modify which tables (for column drop)?", options=table_names)
    drop_cols = st.multiselect("Columns to remove:", options=df.columns.tolist())
    for seg in segments:
        if seg["name"] in to_modify and drop_cols:
            seg["df"] = seg["df"].drop(columns=drop_cols, errors="ignore")

    # 3) Hyperlink per table
    to_link = st.multiselect("Add hyperlink to which tables?", options=table_names)
    link_col = st.selectbox("Hyperlink column:", options=[""] + df.columns.tolist())
    link_row = st.number_input("Row in table to hyperlink (0-based):", min_value=0, value=0)
    link_url = st.text_input("URL for hyperlink:")
    for seg in segments:
        if seg["name"] in to_link and link_col and link_url:
            tbl = seg["df"]
            if link_col in tbl.columns and 0 <= link_row < len(tbl):
                base = str(tbl.at[link_row, link_col])
                tbl.at[link_row, link_col] = f'{base} – <link href="{link_url}">link</link>'

    # Preview each table
    for seg in segments:
        st.subheader(seg["name"])
        st.dataframe(seg["df"].reset_index(drop=True), use_container_width=True)

    # Generate PDF
    title = st.text_input("Proposal Title", os.path.splitext(uploaded.name)[0])
    if st.button("Generate PDF"):
        # Re-append real totals
        for seg in segments:
            seg["df"] = calculate_and_insert_totals(seg["df"])
        pdf_bytes = make_pdf(segments, title)
        st.download_button(
            "📥 Download PDF", data=pdf_bytes, file_name=f"{title}.pdf", mime="application/pdf"
        )


if __name__ == "__main__":
    main()
