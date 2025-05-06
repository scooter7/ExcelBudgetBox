# app.py

import io
import os
import tempfile

import pandas as pd
import requests
import streamlit as st
from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import (Image, Paragraph, SimpleDocTemplate,
                                Spacer, Table, TableStyle)

# URL of the Carnegie logo
LOGO_URL = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"

st.set_page_config(page_title="Proposal → PDF", layout="wide")

def load_dataframe(uploaded_file):
    """Read CSV or Excel into a pandas DataFrame."""
    name = uploaded_file.name.lower()
    if name.endswith((".xls", ".xlsx")):
        return pd.read_excel(uploaded_file)
    else:
        return pd.read_csv(uploaded_file)

def add_strategy_column(df):
    """Split the first line of 'Description' into 'Strategy'."""
    if "Description" not in df.columns:
        st.error("No 'Description' column found.")
        return df

    # Split on first newline
    strategies = []
    descriptions = []
    for val in df["Description"].fillna("").astype(str):
        parts = val.split("\n", 1)
        strategies.append(parts[0])
        descriptions.append(parts[1] if len(parts) > 1 else "")
    df.insert(0, "Strategy", strategies)
    df["Description"] = descriptions
    return df

def replace_est(df):
    """Replace all occurrences of 'Est.' with 'Estimated'."""
    df.columns = [col.replace("Est.", "Estimated") for col in df.columns]
    # also in any string cell
    df = df.applymap(lambda v: v.replace("Est.", "Estimated") if isinstance(v, str) else v)
    return df

def calculate_and_insert_totals(df):
    """Sum up numeric columns and capture last 'Item Total' and 'Estimated Impressions' values."""
    # Define column names
    mcol = "Monthly Amount"
    icol = "Item Total"
    impcol = "Estimated Impressions"
    ccol = "Estimated Conversions"

    # capture totals before dropping
    item_total_val = None
    imp_val = None
    mask_total_rows = df["Description"].str.contains("Total", na=False, case=False)
    if mask_total_rows.any() and icol in df.columns:
        item_total_val = df.loc[mask_total_rows, icol].iat[-1]
    mask_imp = df["Description"].str.contains("Impressions", na=False)
    if mask_imp.any() and impcol in df.columns:
        imp_val = df.loc[mask_imp, impcol].iat[-1]

    # compute sums
    monthly_sum = df[mcol].dropna().apply(pd.to_numeric, errors="coerce").sum()
    conversions_sum = df[ccol].dropna().apply(pd.to_numeric, errors="coerce").sum()

    # drop unwanted rows
    drop_mask = (
        df["Description"].str.contains("Impressions", na=False) |
        df["Description"].str.contains("Conversions", na=False) |
        df["Description"].str.contains("Total", na=False)
    )
    df = df.loc[~drop_mask].reset_index(drop=True)

    # append totals row
    total_row = {c: "" for c in df.columns}
    total_row["Strategy"] = "Total"
    if mcol in df.columns: total_row[mcol] = monthly_sum
    if icol in df.columns and item_total_val is not None: total_row[icol] = item_total_val
    if impcol in df.columns and imp_val is not None: total_row[impcol] = imp_val
    if ccol in df.columns: total_row[ccol] = conversions_sum

    df = df.append(total_row, ignore_index=True)
    return df

def make_pdf(df: pd.DataFrame, title: str) -> io.BytesIO:
    """Build a landscape PDF with logo, title, and styled table."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=18)
    elems = []

    # logo
    resp = requests.get(LOGO_URL)
    logo_img = PILImage.open(io.BytesIO(resp.content))
    # save into a temp file so ReportLab can read it
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    logo_img.save(tmp.name)
    rl_img = Image(tmp.name, width=200, height=50)
    elems.append(rl_img)

    # title
    styles = getSampleStyleSheet()
    elems.append(Spacer(1, 12))
    elems.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
    elems.append(Spacer(1, 12))

    # table data
    data = [list(df.columns)]
    for row in df.itertuples(index=False):
        data.append([str(getattr(row, c)) for c in df.columns])

    table = Table(data, repeatRows=1)
    n_rows = len(data)
    tbl_style = TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        # total row styling
        ("BACKGROUND", (0,n_rows-1), (-1,n_rows-1), colors.lightgrey),
        ("FONTNAME", (0,n_rows-1), (-1,n_rows-1), "Helvetica-Bold"),
    ])
    table.setStyle(tbl_style)
    elems.append(table)

    doc.build(elems)
    buffer.seek(0)
    return buffer

def main():
    st.title("🔄 Proposal Transformer → PDF")
    uploaded = st.file_uploader("Upload Excel or CSV", type=["xls","xlsx","csv"])
    if not uploaded:
        st.info("Please upload a file to begin.")
        return

    df = load_dataframe(uploaded)
    st.markdown("**Preview of uploaded data:**")
    st.dataframe(df.head())

    title_default = os.path.splitext(uploaded.name)[0]
    proposal_title = st.text_input("Proposal Title", value=title_default)

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
            mime="application/pdf"
        )

if __name__ == "__main__":
    main()
