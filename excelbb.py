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
    name = uploaded_file.name.lower()
    if name.endswith((".xls", ".xlsx")):
        return pd.read_excel(uploaded_file)
    else:
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
    # Column names
    mcol = "Monthly Amount"
    icol = "Item Total"
    impcol = "Estimated Impressions"
    ccol = "Estimated Conversions"

    # Capture existing totals
    item_total_val = None
    if icol in df.columns:
        mask = df["Description"].str.contains("Total", na=False, case=False)
        if mask.any():
            item_total_val = df.loc[mask, icol].iat[-1]

    imp_val = None
    if impcol in df.columns:
        mask = df["Description"].str.contains("Impressions", na=False)
        if mask.any():
            imp_val = df.loc[mask, impcol].iat[-1]

    # Sum numeric columns
    monthly_sum = pd.to_numeric(df.get(mcol, pd.Series()), errors="coerce").sum()
    conversions_sum = pd.to_numeric(df.get(ccol, pd.Series()), errors="coerce").sum()

    # Drop old Total/Impressions/Conversions rows
    drop_mask = (
        df["Description"].str.contains("Total|Impressions|Conversions", na=False, case=False)
    )
    df = df.loc[~drop_mask].reset_index(drop=True)

    # Build new total row
    total_row = {c: "" for c in df.columns}
    total_row["Strategy"] = "Total"
    if mcol in df.columns: total_row[mcol] = monthly_sum
    if icol in df.columns and item_total_val is not None: total_row[icol] = item_total_val
    if impcol in df.columns and imp_val is not None: total_row[impcol] = imp_val
    if ccol in df.columns: total_row[ccol] = conversions_sum

    # Append via concat (instead of deprecated DataFrame.append)
    total_df = pd.DataFrame([total_row])
    df = pd.concat([df, total_df], ignore_index=True)
    return df

def make_pdf(df: pd.DataFrame, title: str) -> io.BytesIO:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        rightMargin=30,
        leftMargin=30,
        topMargin=30,
        bottomMargin=18
    )
    elems = []

    # Carnegie logo
    resp = requests.get(LOGO_URL)
    logo_img = PILImage.open(io.BytesIO(resp.content))
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    logo_img.save(tmp.name)
    elems.append(Image(tmp.name, width=200, height=50))

    # Title
    styles = getSampleStyleSheet()
    elems.append(Spacer(1, 12))
    elems.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
    elems.append(Spacer(1, 12))

    # Table
    data = [list(df.columns)] + df.astype(str).values.tolist()
    table = Table(data, repeatRows=1)
    n = len(data)
    style = TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("BACKGROUND", (0,n-1), (-1,n-1), colors.lightgrey),
        ("FONTNAME", (0,n-1), (-1,n-1), "Helvetica-Bold"),
    ])
    table.setStyle(style)
    elems.append(table)

    doc.build(elems)
    buffer.seek(0)
    return buffer

def main():
    st.title("🔄 Proposal Transformer → PDF")
    uploaded = st.file_uploader("Upload Excel or CSV", type=["xls","xlsx","csv"])
    if not uploaded:
        return

    df = load_dataframe(uploaded)
    st.dataframe(df.head())

    proposal_title = st.text_input("Proposal Title", os.path.splitext(uploaded.name)[0])

    if st.button("Generate PDF"):
        with st.spinner("Processing…"):
            df1 = add_strategy_column(df.copy())
            df2 = replace_est(df1)
            df3 = calculate_and_insert_totals(df2)
            pdf_bytes = make_pdf(df3, proposal_title)

        st.download_button(
            "📥 Download PDF",
            data=pdf_bytes,
            file_name=f"{proposal_title}.pdf",
            mime="application/pdf"
        )

if __name__ == "__main__":
    main()
