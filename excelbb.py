# app.py

import io
import os
import tempfile

import pandas as pd
import requests
import streamlit as st
from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, ledger
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import (
    Image,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# Carnegie logo URL
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

    # Compute sums
    monthly_sum = pd.to_numeric(df.get(mcol, []), errors="coerce").sum()
    conversions_sum = pd.to_numeric(df.get(ccol, []), errors="coerce").sum()

    # Drop old Total/Impressions/Conversions rows
    drop_mask = (
        df["Description"]
          .str.contains("Total|Impressions|Conversions", na=False, case=False)
    )
    df = df.loc[~drop_mask].reset_index(drop=True)

    # Build new total row
    total_row = {c: "" for c in df.columns}
    total_row["Strategy"] = "Total"
    if mcol in df.columns: total_row[mcol] = monthly_sum
    if icol in df.columns and item_total_val is not None: total_row[icol] = item_total_val
    if impcol in df.columns and imp_val is not None: total_row[impcol] = imp_val
    if ccol in df.columns: total_row[ccol] = conversions_sum

    # Append via concat (avoids deprecated .append)
    total_df = pd.DataFrame([total_row])
    df = pd.concat([df, total_df], ignore_index=True)
    return df

def make_pdf(df: pd.DataFrame, title: str) -> io.BytesIO:
    # 1) Prepare df: parse & format dates, then blank out NaN/NaT
    pdf_df = df.copy()
    for col in pdf_df.columns:
        if "date" in col.lower():
            pdf_df[col] = pd.to_datetime(pdf_df[col], errors="coerce")
        if pd.api.types.is_datetime64_any_dtype(pdf_df[col]):
            pdf_df[col] = pdf_df[col].dt.strftime("%m/%d/%Y")
    pdf_df = pdf_df.fillna("")

    # 2) Build the PDF
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(ledger),
        rightMargin=30, leftMargin=30,
        topMargin=30, bottomMargin=18,
    )
    elems = []

    # Logo
    resp = requests.get(LOGO_URL)
    logo_img = PILImage.open(io.BytesIO(resp.content))
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    logo_img.save(tmp.name)
    elems.append(Image(tmp.name, width=200, height=50))
    elems.append(Spacer(1, 12))

    # Title
    styles = getSampleStyleSheet()
    elems.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
    elems.append(Spacer(1, 12))

    # Table
    data = [list(pdf_df.columns)] + pdf_df.astype(str).values.tolist()
    col_count = len(pdf_df.columns)
    table = Table(
        data,
        colWidths=[doc.width / col_count] * col_count,
        repeatRows=1
    )
    n = len(data)
    style = TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
        ("
