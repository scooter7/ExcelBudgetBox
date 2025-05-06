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


def load_dataframe(uploaded_file):
    fn = uploaded_file.name.lower()
    return (
        pd.read_excel(uploaded_file)
        if fn.endswith((".xls", ".xlsx"))
        else pd.read_csv(uploaded_file)
    )


def add_strategy_column(df):
    if "Description" not in df.columns:
        st.error("No 'Description' column found.")
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

    itot = None
    if icol in df.columns and "Total" in df["Strategy"].values:
        itot = df.loc[df["Strategy"] == "Total", icol].iat[-1]
    impv = None
    if impcol in df.columns and "Est. Conversions" in df["Strategy"].values:
        impv = df.loc[df["Strategy"] == "Est. Conversions", impcol].iat[-1]

    monthly_sum = pd.to_numeric(df.get(mcol, []), errors="coerce").sum()
    conv_sum = pd.to_numeric(df.get(ccol, []), errors="coerce").sum()

    drop_exact = {
        "Total",
        "Est. Conversions",
        "Estimated Conversions",
        "Est. Impressions",
        "Estimated Impressions",
    }
    df_clean = df.loc[~df["Strategy"].isin(drop_exact)].reset_index(drop=True)

    total_row = {c: "" for c in df_clean.columns}
    total_row["Strategy"] = "Total"
    total_row[mcol] = monthly_sum if mcol in total_row else ""
    if icol in total_row and itot is not None:
        total_row[icol] = itot
    if impcol in total_row and impv is not None:
        total_row[impcol] = impv
    total_row[ccol] = conv_sum if ccol in total_row else ""

    return pd.concat([df_clean, pd.DataFrame([total_row])], ignore_index=True)


def make_pdf(df: pd.DataFrame, title: str) -> io.BytesIO:
    pdf_df = df.copy()
    for col in pdf_df.columns:
        if "date" in col.lower():
            pdf_df[col] = pd.to_datetime(pdf_df[col], errors="coerce")
        if pd.api.types.is_datetime64_any_dtype(pdf_df[col]):
            pdf_df[col] = pdf_df[col].dt.strftime("%m/%d/%Y")
    pdf_df = pdf_df.fillna("")

    for col in ("Monthly Amount", "Item Total"):
        if col in pdf_df.columns:
            pdf_df[col] = (
                pd.to_numeric(pdf_df[col], errors="coerce")
                .fillna(0)
                .map(lambda x: f"${x:,.0f}")
            )

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

    # Header style: DMSerif
    header_style = ParagraphStyle(
        "hdr",
        parent=styles["BodyText"],
        fontName="DMSerif",
        fontSize=8,
        leading=9,
        alignment=1,
    )
    # Body style: Barlow
    wrap_style = ParagraphStyle(
        "wrap",
        parent=styles["BodyText"],
        fontName="Barlow",
        fontSize=7,
        leading=8,
        alignment=0,
    )

    # Logo ×3
    resp = requests.get(LOGO_URL)
    logo_img = PILImage.open(io.BytesIO(resp.content))
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    logo_img.save(tmp.name)
    elems.append(Image(tmp.name, width=4.5 * inch, height=1.5 * inch))
    elems.append(Spacer(1, 12))

    elems.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
    elems.append(Spacer(1, 12))

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

    widths = [0.08, 0.30, 0.06, 0.08, 0.08, 0.12, 0.12, 0.06, 0.10]
    col_widths = [doc.width * w for w in widths]

    table = Table(data, colWidths=col_widths, repeatRows=1)
    tbl_style = TableStyle(
        [
            ("GRID", (0, 0), (-1, -1), 0.4, colors.black),
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("FONTNAME", (0, 0), (-1, 0), "DMSerif"),
            ("FONTNAME", (0, 1), (-1, -1), "Barlow"),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BACKGROUND", (0, len(data) - 1), (-1, len(data) - 1), colors.lightgrey),
            ("FONTNAME", (0, len(data) - 1), (-1, len(data) - 1), "DMSerif"),  # total row in DMSerif
        ]
    )
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

    default_title = os.path.splitext(uploaded.name)[0]
    proposal_title = st.text_input("Proposal Title", default_title)

    if st.button("Generate PDF"):
        with st.spinner("Rendering PDF…"):
            df1 = add_strategy_column(df.copy())
            df2 = replace_est(df1)
            df3 = calculate_and_insert_totals(df2)
            pdf_bytes = make_pdf(df3, proposal_title)

        st.download_button(
            "📥 Download PDF",
            data=pdf_bytes,
            file_name=f"{proposal_title}.pdf",
            mime="application/pdf",
        )

if __name__ == "__main__":
    main()
