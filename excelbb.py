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

LOGO_URL = (
    "https://www.carnegiehighered.com/wp-content/uploads/2021/11/"
    "Twitter-Image-2-2021.png"
)

# Raw URLs for the fonts
FONTS = {
    "Barlow": "https://raw.githubusercontent.com/scooter7/ExcelBudgetBox/main/fonts/Barlow-Regular.ttf",
    "DMSerif": "https://raw.githubusercontent.com/scooter7/ExcelBudgetBox/main/fonts/DMSerifDisplay-Regular.ttf",
}

def register_fonts():
    for name, url in FONTS.items():
        r = requests.get(url)
        path = tempfile.NamedTemporaryFile(delete=False, suffix=".ttf").name
        with open(path, "wb") as f:
            f.write(r.content)
        pdfmetrics.registerFont(TTFont(name, path))

register_fonts()

st.set_page_config(page_title="Proposal → PDF", layout="wide")


def load_and_prepare_dataframe(uploaded_file):
    """Load Excel/CSV and rename first column to 'Service'."""
    fn = uploaded_file.name.lower()
    df = pd.read_excel(uploaded_file) if fn.endswith((".xls", ".xlsx")) else pd.read_csv(uploaded_file)
    # rename whatever is in column A to 'Service'
    first_col = df.columns[0]
    if first_col != "Service":
        df = df.rename(columns={first_col: "Service"})
    return df


def transform_service_column(df):
    """Clean the 'Service' column: remove leading 'XX:', slash+suffix, and '(Markup)'."""
    def clean(s):
        s = str(s)
        # remove first two chars + colon
        s = re.sub(r'^..:', '', s)
        # strip off '/' and everything after
        s = s.split('/', 1)[0]
        # remove '(Markup)' case-insensitive
        s = re.sub(r'\(\s*Markup\s*\)', '', s, flags=re.IGNORECASE)
        return s.strip()
    df["Service"] = df["Service"].fillna("").apply(clean)
    return df


def replace_est(df):
    """Replace 'Est.' with 'Estimated' everywhere."""
    df.columns = [c.replace("Est.", "Estimated") for c in df.columns]
    return df.applymap(lambda v: v.replace("Est.", "Estimated") if isinstance(v, str) else v)


def calculate_and_insert_totals(df):
    """Sum columns, capture original footer values, drop legacy footer rows, append new Total row."""
    mcol = "Monthly Amount"
    icol = "Item Total"
    impcol = "Estimated Impressions"
    ccol = "Estimated Conversions"

    # capture original footer values
    itot = None
    if icol in df.columns and "Total" in df["Service"].values:
        itot = df.loc[df["Service"] == "Total", icol].iat[-1]
    impv = None
    if impcol in df.columns and "Est. Conversions" in df["Service"].values:
        impv = df.loc[df["Service"] == "Est. Conversions", impcol].iat[-1]

    # sums
    monthly_sum = pd.to_numeric(df.get(mcol, []), errors="coerce").sum()
    conv_sum = pd.to_numeric(df.get(ccol, []), errors="coerce").sum()

    # drop exact legacy footer rows
    drop_exact = {
        "Total", "Est. Conversions", "Estimated Conversions",
        "Est. Impressions", "Estimated Impressions"
    }
    df_clean = df.loc[~df["Service"].isin(drop_exact)].reset_index(drop=True)

    # build new Total row
    total_row = {c: "" for c in df_clean.columns}
    total_row["Service"] = "Total"
    if mcol in total_row:
        total_row[mcol] = monthly_sum
    if icol in total_row and itot is not None:
        total_row[icol] = itot
    if impcol in total_row and impv is not None:
        total_row[impcol] = impv
    if ccol in total_row:
        total_row[ccol] = conv_sum

    return pd.concat([df_clean, pd.DataFrame([total_row])], ignore_index=True)


def make_pdf(df: pd.DataFrame, title: str,
             hyperlink_col: str, hyperlink_row: int, hyperlink_url: str) -> io.BytesIO:
    """Generate the landscape 11×17 PDF with logo, title, table, and optional hyperlink."""
    # format dates
    pdf_df = df.copy()
    for col in pdf_df.columns:
        if "date" in col.lower():
            pdf_df[col] = pd.to_datetime(pdf_df[col], errors="coerce")
        if pd.api.types.is_datetime64_any_dtype(pdf_df[col]):
            pdf_df[col] = pdf_df[col].dt.strftime("%m/%d/%Y")
    pdf_df = pdf_df.fillna("")

    # currency formatting
    for col in ("Monthly Amount", "Item Total"):
        if col in pdf_df.columns:
            pdf_df[col] = (
                pd.to_numeric(pdf_df[col], errors="coerce")
                .fillna(0)
                .map(lambda x: f"${x:,.0f}")
            )

    # insert hyperlink if provided
    if hyperlink_col in pdf_df.columns and hyperlink_row > 1:
        idx = hyperlink_row - 2
        if 0 <= idx < len(pdf_df):
            text = str(pdf_df.at[idx, hyperlink_col])
            link_html = f'<a href="{hyperlink_url}">{text}-link</a>'
            pdf_df.at[idx, hyperlink_col] = link_html

    # build PDF
    buffer = io.BytesIO()
    tw, th = 17 * inch, 11 * inch
    doc = SimpleDocTemplate(
        buffer,
        pagesize=(tw, th),
        leftMargin=0.5 * inch,
        rightMargin=0.5 * inch,
        topMargin=0.5 * inch,
        bottomMargin=0.5 * inch,
    )

    elems = []
    styles = getSampleStyleSheet()

    # header style: DMSerif
    header_style = ParagraphStyle(
        "hdr",
        parent=styles["BodyText"],
        fontName="DMSerif",
        fontSize=8,
        leading=9,
        alignment=1,
    )
    # body style: Barlow
    body_style = ParagraphStyle(
        "body",
        parent=styles["BodyText"],
        fontName="Barlow",
        fontSize=7,
        leading=8,
        alignment=0,
    )

    # logo ×3
    resp = requests.get(LOGO_URL)
    logo_img = PILImage.open(io.BytesIO(resp.content))
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    logo_img.save(tmp.name)
    elems.append(Image(tmp.name, width=4.5 * inch, height=1.5 * inch))
    elems.append(Spacer(1, 12))

    elems.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
    elems.append(Spacer(1, 12))

    # build table data
    header_cells = [Paragraph(col, header_style) for col in pdf_df.columns]
    data = [header_cells]
    for row in pdf_df.itertuples(index=False):
        cells = []
        for col, val in zip(pdf_df.columns, row):
            txt = "" if val is None else str(val)
            if txt.startswith("<a "):
                cells.append(Paragraph(txt, body_style))
            elif col in ("Service", "Description", "Notes"):
                cells.append(Paragraph(txt, body_style))
            else:
                cells.append(txt)
        data.append(cells)

    # column widths proportionally
    widths = [0.12, 0.30, 0.06, 0.08, 0.08, 0.12, 0.12, 0.06, 0.06]
    col_widths = [doc.width * w for w in widths]

    table = Table(data, colWidths=col_widths, repeatRows=1)
    tbl_style = TableStyle([
        ("GRID", (0,0), (-1,-1), 0.4, colors.black),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("FONTNAME", (0,0), (-1,0), "DMSerif"),
        ("FONTNAME", (0,1), (-1,-1), "Barlow"),
        ("FONTSIZE", (0,0), (-1,-1), 7),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("BACKGROUND", (0,len(data)-1), (-1,len(data)-1), colors.lightgrey),
        ("FONTNAME", (0,len(data)-1), (-1,len(data)-1), "DMSerif"),
    ])
    table.setStyle(tbl_style)
    elems.append(table)

    doc.build(elems)
    buffer.seek(0)
    return buffer


def main():
    st.title("🔄 Proposal → PDF")

    uploaded = st.file_uploader("Upload Excel or CSV", type=["xls","xlsx","csv"])
    if not uploaded:
        return

    # load and rename first column to Service
    df = load_and_prepare_dataframe(uploaded)
    st.dataframe(df.head())

    # transform Service entries
    df = transform_service_column(df)

    # allow dropping columns
    to_drop = st.multiselect("Columns to remove:", options=df.columns.tolist())
    if to_drop:
        df = df.drop(columns=to_drop)

    # hyperlink inputs
    hyperlink_col = st.selectbox("Column for hyperlink:", options=[""] + df.columns.tolist())
    hyperlink_row = st.number_input("Excel row to hyperlink (1=header):", min_value=1, max_value=len(df)+1, value=2)
    hyperlink_url = st.text_input("URL for hyperlink:")

    proposal_title = st.text_input("Proposal Title", os.path.splitext(uploaded.name)[0])

    if st.button("Generate PDF"):
        with st.spinner("Rendering PDF…"):
            df2 = replace_est(df.copy())
            df3 = calculate_and_insert_totals(df2)
            pdf_bytes = make_pdf(df3, proposal_title, hyperlink_col, hyperlink_row, hyperlink_url)

        st.download_button("📥 Download PDF", data=pdf_bytes, file_name=f"{proposal_title}.pdf", mime="application/pdf")


if __name__ == "__main__":
    main()
