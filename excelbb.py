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
from reportlab.lib.styles import getSampleStyleSheet
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
    name = uploaded_file.name.lower()
    if name.endswith((".xls", ".xlsx")):
        return pd.read_excel(uploaded_file)
    return pd.read_csv(uploaded_file)


def add_strategy_column(df):
    if "Description" not in df.columns:
        st.error("No 'Description' column found.")
        return df
    strategies, descriptions = [], []
    for val in df["Description"].fillna("").astype(str):
        parts = val.split("\n", 1)
        strategies.append(parts[0])
        descriptions.append(parts[1] if len(parts) > 1 else "")
    df.insert(0, "Strategy", strategies)
    df["Description"] = descriptions
    return df


def replace_est(df):
    df.columns = [col.replace("Est.", "Estimated") for col in df.columns]
    return df.applymap(lambda v: v.replace("Est.", "Estimated") if isinstance(v, str) else v)


def calculate_and_insert_totals(df):
    mcol = "Monthly Amount"
    icol = "Item Total"
    impcol = "Estimated Impressions"
    ccol = "Estimated Conversions"

    # capture last totals
    item_total_val = (
        df.loc[df["Description"].str.contains("Total", na=False), icol]
        .iat[-1]
        if icol in df.columns and df["Description"].str.contains("Total", na=False).any()
        else None
    )
    imp_val = (
        df.loc[df["Description"].str.contains("Impressions", na=False), impcol]
        .iat[-1]
        if impcol in df.columns and df["Description"].str.contains("Impressions", na=False).any()
        else None
    )

    monthly_sum = pd.to_numeric(df.get(mcol, []), errors="coerce").sum()
    conversions_sum = pd.to_numeric(df.get(ccol, []), errors="coerce").sum()

    # drop old total/imp/conversion rows
    mask = df["Description"].str.contains("Total|Impressions|Conversions", na=False, case=False)
    df = df.loc[~mask].reset_index(drop=True)

    # build & append new total row
    total_row = {c: "" for c in df.columns}
    total_row["Strategy"] = "Total"
    if mcol in df.columns: total_row[mcol] = monthly_sum
    if icol in df.columns and item_total_val is not None: total_row[icol] = item_total_val
    if impcol in df.columns and imp_val is not None: total_row[impcol] = imp_val
    if ccol in df.columns: total_row[ccol] = conversions_sum

    df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
    return df


def make_pdf(df: pd.DataFrame, title: str) -> io.BytesIO:
    # 1) Format dates, blank out NaT/nan
    pdf_df = df.copy()
    for col in pdf_df.columns:
        if "date" in col.lower():
            pdf_df[col] = pd.to_datetime(pdf_df[col], errors="coerce")
        if pd.api.types.is_datetime64_any_dtype(pdf_df[col]):
            pdf_df[col] = pdf_df[col].dt.strftime("%m/%d/%Y")
    pdf_df = pdf_df.fillna("")

    # 2) Currency formatting
    for col in ("Monthly Amount", "Item Total"):
        if col in pdf_df.columns:
            pdf_df[col] = (
                pd.to_numeric(pdf_df[col], errors="coerce")
                .fillna(0)
                .map(lambda x: f"${x:,.0f}")
            )

    # 3) Build PDF
    buffer = io.BytesIO()
    page_size = (17 * inch, 11 * inch)
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(page_size),
        rightMargin=30,
        leftMargin=30,
        topMargin=30,
        bottomMargin=18,
    )

    elems = []
    styles = getSampleStyleSheet()
    normal_wrapped = styles["BodyText"]
    normal_wrapped.fontSize = 8
    normal_wrapped.leading = 10

    # logo
    resp = requests.get(LOGO_URL)
    logo_img = PILImage.open(io.BytesIO(resp.content))
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    logo_img.save(tmp.name)
    elems.append(Image(tmp.name, width=200, height=50))
    elems.append(Spacer(1, 12))

    # title
    elems.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
    elems.append(Spacer(1, 12))

    # build table data with Paragraphs for long text
    data = [list(pdf_df.columns)]
    for row in pdf_df.itertuples(index=False):
        cells = []
        for col, val in zip(pdf_df.columns, row):
            txt = "" if val is None else str(val)
            if col in ("Description", "Notes"):
                cells.append(Paragraph(txt, normal_wrapped))
            else:
                cells.append(txt)
        data.append(cells)

    # column widths
    total_w = doc.width
    col_widths = []
    for col in pdf_df.columns:
        if col == "Description":
            col_widths.append(total_w * 0.35)
        elif col == "Strategy":
            col_widths.append(total_w * 0.10)
        elif col in ("Term (Months)", "Estimated Conversions"):
            col_widths.append(total_w * 0.08)
        elif col in ("Start Date", "End Date"):
            col_widths.append(total_w * 0.10)
        elif col in ("Monthly Amount", "Item Total"):
            col_widths.append(total_w * 0.10)
        else:  # Notes or any other
            col_widths.append(total_w * 0.17)

    table = Table(data, colWidths=col_widths, repeatRows=1)

    tbl_style = TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, len(data) - 1), (-1, len(data) - 1), colors.lightgrey),
        ("FONTNAME", (0, len(data) - 1), (-1, len(data) - 1), "Helvetica-Bold"),
    ])
    table.setStyle(tbl_style)
    elems.append(table)

    doc.build(elems)
    buffer.seek(0)
    return buffer


def main():
    st.title("🔄 Proposal Transformer → PDF")
    uploaded = st.file_uploader("Upload Excel or CSV", type=["xls", "xlsx", "csv"])
    if not uploaded:
        st.info("Please upload a file to begin.")
        return

    df = load_dataframe(uploaded)
    st.markdown("**Data preview:**")
    st.dataframe(df.head())

    default_title = os.path.splitext(uploaded.name)[0]
    proposal_title = st.text_input("Proposal Title", default_title)

    if st.button("Generate PDF"):
        with st.spinner("Processing…"):
            df1 = add_strategy_column(df.copy())
            df2 = replace_est(df1)
            df3 = calculate_and_insert_totals(df2)
            pdf_bytes = make_pdf(df3, proposal_title)

        st.success("Done! Download below:")
        st.download_button(
            "📥 Download PDF",
            data=pdf_bytes,
            file_name=f"{proposal_title}.pdf",
            mime="application/pdf",
        )


if __name__ == "__main__":
    main()
