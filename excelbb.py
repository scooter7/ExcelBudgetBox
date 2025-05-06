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
    return (
        pd.read_excel(uploaded_file)
        if fn.endswith((".xls", ".xlsx"))
        else pd.read_csv(uploaded_file)
    )


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

    # capture existing values
    itot = (
        df.loc[df["Description"].str.contains("Total", case=False, na=False), icol]
        .iat[-1]
        if icol in df.columns and df["Description"].str.contains("Total", na=False, case=False).any()
        else None
    )
    impv = (
        df.loc[df["Description"].str.contains("Impressions", na=False), impcol]
        .iat[-1]
        if impcol in df.columns and df["Description"].str.contains("Impressions", na=False).any()
        else None
    )

    ms = pd.to_numeric(df.get(mcol, []), errors="coerce").sum()
    cs = pd.to_numeric(df.get(ccol, []), errors="coerce").sum()

    # drop any row where Strategy or Description mentions Total, Impressions, or Conversions
    drop_re = r"Total|Impressions|Conversions"
    mask = (
        df["Strategy"].str.contains(drop_re, case=False, na=False)
        | df["Description"].str.contains(drop_re, case=False, na=False)
    )
    df = df.loc[~mask].reset_index(drop=True)

    # build new total row
    row = {c: "" for c in df.columns}
    row["Strategy"] = "Total"
    if mcol in row: row[mcol] = ms
    if icol in row and itot is not None: row[icol] = itot
    if impcol in row and impv is not None: row[impcol] = impv
    if ccol in row: row[ccol] = cs

    return pd.concat([df, pd.DataFrame([row])], ignore_index=True)


def make_pdf(df: pd.DataFrame, title: str) -> io.BytesIO:
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

    # set up true 11"x17" landscape
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

    # header style (for wrapping column names)
    header_style = ParagraphStyle(
        "hdr",
        parent=styles["BodyText"],
        fontSize=8,
        leading=9,
        alignment=1,        # center
        spaceAfter=4,
        spaceBefore=4,
        fontName="Helvetica-Bold"
    )

    # body wrap style
    wrap_style = ParagraphStyle(
        "wrap",
        parent=styles["BodyText"],
        fontSize=7,
        leading=8,
        alignment=0,        # left
    )

    # logo + title
    resp = requests.get(LOGO_URL)
    img = PILImage.open(io.BytesIO(resp.content))
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    img.save(tmp.name)
    elems.append(Image(tmp.name, width=1.5 * inch, height=0.5 * inch))
    elems.append(Spacer(1, 12))
    elems.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
    elems.append(Spacer(1, 12))

    # build table data
    # 1) header row as Paragraphs
    header_cells = [Paragraph(col, header_style) for col in pdf_df.columns]
    data = [header_cells]

    # 2) body rows
    for row in pdf_df.itertuples(index=False):
        cells = []
        for col, val in zip(pdf_df.columns, row):
            text = "" if val is None else str(val)
            if col in ("Strategy", "Description", "Notes"):
                cells.append(Paragraph(text, wrap_style))
            else:
                cells.append(text)
        data.append(cells)

    # column widths summing to doc.width
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
            "📥 Download PDF", data=pdf, file_name=f"{title}.pdf", mime="application/pdf"
        )


if __name__ == "__main__":
    main()
