# app.py

import io
import os
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

LOGO_URL = (
    "https://www.carnegiehighered.com/wp-content/uploads/2021/11/"
    "Twitter-Image-2-2021.png"
)

st.set_page_config(page_title="Proposal → PDF", layout="wide")


def load_dataframe(uploaded_file):
    fn = uploaded_file.name.lower()
    if fn.endswith((".xls", ".xlsx")):
        return pd.read_excel(uploaded_file)
    return pd.read_csv(uploaded_file)


def add_strategy_column(df):
    if "Description" not in df.columns:
        st.error("No 'Description' column")
        return df
    strat, desc = [], []
    for v in df["Description"].fillna("").astype(str):
        parts = v.split("\n", 1)
        strat.append(parts[0])
        desc.append(parts[1] if len(parts) > 1 else "")
    df.insert(0, "Strategy", strat)
    df["Description"] = desc
    return df


def replace_est(df):
    df.columns = [c.replace("Est.", "Estimated") for c in df.columns]
    return df.applymap(lambda v: v.replace("Est.", "Estimated") if isinstance(v, str) else v)


def calculate_and_insert_totals(df):
    mcol = "Monthly Amount"
    icol = "Item Total"
    impcol = "Estimated Impressions"
    ccol = "Estimated Conversions"

    # Capture original totals from the Strategy column
    itot = None
    if icol in df.columns:
        mask_tot = df["Strategy"].astype(str).str.contains("Total", na=False, case=False)
        if mask_tot.any():
            itot = df.loc[mask_tot, icol].iat[-1]

    impv = None
    if impcol in df.columns:
        mask_imp = df["Strategy"].astype(str).str.contains("Impressions", na=False, case=False)
        if mask_imp.any():
            impv = df.loc[mask_imp, impcol].iat[-1]

    # Sum numeric columns
    ms = pd.to_numeric(df.get(mcol, []), errors="coerce").sum()
    cs = pd.to_numeric(df.get(ccol, []), errors="coerce").sum()

    # Drop any old Total/Impressions/Conversions rows in either Strategy or Description
    drop_re = r"Total|Impressions|Conversions"
    mask = (
        df["Strategy"].astype(str).str.contains(drop_re, na=False, case=False)
        | df["Description"].astype(str).str.contains(drop_re, na=False, case=False)
    )
    df = df.loc[~mask].reset_index(drop=True)

    # Build new Total row
    total_row = {c: "" for c in df.columns}
    total_row["Strategy"] = "Total"
    if mcol in total_row:
        total_row[mcol] = ms
    if icol in total_row and itot is not None:
        total_row[icol] = itot
    if impcol in total_row and impv is not None:
        total_row[impcol] = impv
    if ccol in total_row:
        total_row[ccol] = cs

    return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)


def make_pdf(df: pd.DataFrame, title: str) -> io.BytesIO:
    # Format dates and blank out NaT/nan
    pdf_df = df.copy()
    for col in pdf_df.columns:
        if "date" in col.lower():
            pdf_df[col] = pd.to_datetime(pdf_df[col], errors="coerce")
        if pd.api.types.is_datetime64_any_dtype(pdf_df[col]):
            pdf_df[col] = pdf_df[col].dt.strftime("%m/%d/%Y")
    pdf_df = pdf_df.fillna("")

    # Currency formatting
    for col in ("Monthly Amount", "Item Total"):
        if col in pdf_df.columns:
            pdf_df[col] = (
                pd.to_numeric(pdf_df[col], errors="coerce")
                .fillna(0)
                .map(lambda x: f"${x:,.0f}")
            )

    # Setup true 11"x17" landscape
    buf = io.BytesIO()
    tw, th = 17 * inch, 11 * inch
    doc = SimpleDocTemplate(
        buf,
        pagesize=(tw, th),
        leftMargin=0.5 * inch,
        rightMargin=0.5 * inch,
        topMargin=0.5 * inch,
        bottomMargin=0.5 * inch,
    )

    elems = []
    styles = getSampleStyleSheet()

    header_style = ParagraphStyle(
        "hdr",
        parent=styles["BodyText"],
        fontSize=8,
        leading=9,
        alignment=1,  # center
        fontName="Helvetica-Bold",
    )

    wrap_style = ParagraphStyle(
        "wrap",
        parent=styles["BodyText"],
        fontSize=7,
        leading=8,
        alignment=0,  # left
    )

    # Logo + title
    resp = requests.get(LOGO_URL)
    img = PILImage.open(io.BytesIO(resp.content))
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    img.save(tmp.name)
    elems.append(Image(tmp.name, width=1.5 * inch, height=0.5 * inch))
    elems.append(Spacer(1, 12))
    elems.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
    elems.append(Spacer(1, 12))

    # Build table data
    header_cells = [Paragraph(col, header_style) for col in pdf_df.columns]
    data = [header_cells]
    for row in pdf_df.itertuples(index=False):
        cells = []
        for col, val in zip(pdf_df.columns, row):
            txt = "" if val is None else str(val)
            if col in ("Strategy", "Description", "Notes"):
                cells.append(Paragraph(txt, wrap_style))
            else:
                cells.append(txt)
        data.append(cells)

    # Column widths
    widths = [0.08, 0.30, 0.06, 0.08, 0.08, 0.12, 0.12, 0.06, 0.10]
    col_widths = [doc.width * w for w in widths]

    table = Table(data, colWidths=col_widths, repeatRows=1)
    tbl_style = TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.4, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, len(data)-1), (-1, len(data)-1), colors.lightgrey),
        ("FONTNAME", (0, len(data)-1), (-1, len(data)-1), "Helvetica-Bold"),
    ])
    table.setStyle(tbl_style)
    elems.append(table)

    doc.build(elems)
    buf.seek(0)
    return buf


def main():
    st.title("🔄 Proposal → PDF")
    uploaded = st.file_uploader("Upload Excel or CSV", type=["xls", "xlsx", "csv"])
    if not uploaded:
        return

    df = load_dataframe(uploaded)
    st.dataframe(df.head())

    default = os.path.splitext(uploaded.name)[0]
    title = st.text_input("Proposal Title", default)

    if st.button("Generate PDF"):
        with st.spinner("Rendering PDF…"):
            d1 = add_strategy_column(df.copy())
            d2 = replace_est(d1)
            d3 = calculate_and_insert_totals(d2)
            pdf = make_pdf(d3, title)

        st.download_button(
            "📥 Download PDF",
            data=pdf,
            file_name=f"{title}.pdf",
            mime="application/pdf",
        )


if __name__ == "__main__":
    main()
