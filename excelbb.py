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

# --- Logo & Fonts ---
LOGO_URL = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
FONTS = {
    "Barlow":  "https://raw.githubusercontent.com/scooter7/ExcelBudgetBox/main/fonts/Barlow-Regular.ttf",
    "DMSerif": "https://raw.githubusercontent.com/scooter7/ExcelBudgetBox/main/fonts/DMSerifDisplay-Regular.ttf",
}
for name, url in FONTS.items():
    r = requests.get(url)
    path = tempfile.NamedTemporaryFile(delete=False, suffix=".ttf").name
    open(path, "wb").write(r.content)
    pdfmetrics.registerFont(TTFont(name, path))

st.set_page_config(page_title="Proposal → PDF", layout="wide")


def load_and_prepare_dataframe(uploaded_file):
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
    def clean(s):
        s = str(s)
        s = re.sub(r'^..:', '', s)
        s = s.split('/', 1)[0]
        s = re.sub(r'\(.*?\)', '', s)
        return s.strip()
    df["Service"] = df["Service"].fillna("").apply(clean)
    return df


def replace_est(df):
    df.columns = [c.replace("Est.", "Estimated").strip() for c in df.columns]
    return df.applymap(lambda v: v.replace("Est.", "Estimated") if isinstance(v, str) else v)


def split_tables(df):
    segments = []
    start = 0
    for i, svc in enumerate(df["Service"]):
        if str(svc).strip().lower() == "total":
            seg = df.iloc[start : i + 1].reset_index(drop=True)
            name = next(
                (x for x in seg["Service"] if x and x.strip().lower() != "service"),
                f"Table{len(segments)+1}"
            )
            segments.append({"name": name, "df": seg})
            start = i + 1
    if start < len(df):
        seg = df.iloc[start:].reset_index(drop=True)
        name = next(
            (x for x in seg["Service"] if x and x.strip().lower() != "service"),
            f"Table{len(segments)+1}"
        )
        segments.append({"name": name, "df": seg})
    return segments


def calculate_and_insert_totals(seg_df):
    df = seg_df.copy()
    mask = df["Service"].str.strip().str.lower() == "total"
    orig = df.loc[mask].iloc[0] if mask.any() else None
    drop_keys = {
        "total", "est. conversions", "estimated conversions",
        "est. impressions", "estimated impressions",
    }
    df = df.loc[~df["Service"].str.strip().str.lower().isin(drop_keys)].reset_index(drop=True)
    total = {c: "" for c in df.columns}
    total["Service"] = "Total"
    if orig is not None:
        for c in df.columns:
            v = orig.get(c)
            if pd.notna(v):
                total[c] = v
    return pd.concat([df, pd.DataFrame([total])], ignore_index=True)


def make_pdf(segments, title):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=(17 * inch, 11 * inch),
        leftMargin=0.5 * inch, rightMargin=0.5 * inch,
        topMargin=0.5 * inch, bottomMargin=0.5 * inch,
    )
    styles = getSampleStyleSheet()
    hdr = ParagraphStyle("hdr", parent=styles["BodyText"], fontName="DMSerif", fontSize=8, leading=9, alignment=1)
    bod = ParagraphStyle("bod", parent=styles["BodyText"], fontName="Barlow", fontSize=7, leading=8, alignment=0)

    elems = []
    # Logo
    r = requests.get(LOGO_URL)
    logo_img = PILImage.open(io.BytesIO(r.content))
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    logo_img.save(tmp.name)
    elems.append(Image(tmp.name, width=4.5 * inch, height=1.5 * inch))
    elems.append(Spacer(1, 12))
    elems.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
    elems.append(Spacer(1, 12))

    for seg in segments:
        df = seg["df"]
        df = df[df["Service"].notna() & df["Service"].str.strip().astype(bool)]
        df = df.loc[
            ~(
                (df["Service"].str.strip().str.lower() == "service")
                & (df.get("Description", "").str.strip().str.lower() == "description")
            )
        ].reset_index(drop=True)

        df = calculate_and_insert_totals(df)

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

        header_cells = [Paragraph(c, hdr) for c in df.columns]
        data = [header_cells]
        for row in df.itertuples(index=False):
            cells = []
            for c, v in zip(df.columns, row):
                txt = "" if v is None else str(v)
                if "<a href" in txt or c in ("Service", "Description", "Notes"):
                    cells.append(Paragraph(txt, bod))
                else:
                    cells.append(txt)
            data.append(cells)

        col_ws = [0.12, 0.30, 0.06, 0.08, 0.08, 0.12, 0.12, 0.06, 0.06]
        cw = [doc.width * w for w in col_ws]
        tbl = Table(data, colWidths=cw, repeatRows=1)
        tbl.setStyle(
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
        elems.append(tbl)
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
    segments = split_tables(df)

    # Inline edit rows & columns per table
    for seg in segments:
        # remove Estimated Conversions rows
        df_seg = seg["df"].loc[
            ~seg["df"]["Service"].str.strip().str.lower().eq("estimated conversions")
        ].reset_index(drop=True)

        st.markdown(f"**Edit table: {seg['name']}**")
        keep_cols = st.multiselect(
            "Columns to show/edit:",
            options=df_seg.columns.tolist(),
            default=df_seg.columns.tolist(),
            key=f"cols_{seg['name']}"
        )
        edited = st.data_editor(
            df_seg[keep_cols],
            num_rows="dynamic",
            use_container_width=True,
            key=f"editor_{seg['name']}"
        )
        seg["df"] = edited.copy()

    # Hyperlink per table
    table_names = [s["name"] for s in segments]
    link_tables = st.multiselect("Add hyperlink to which tables?", options=table_names)
    link_col = st.selectbox("Hyperlink column:", options=[""] + df.columns.tolist())
    link_row = st.number_input("Row to hyperlink (0-based):", min_value=0, value=0)
    link_url = st.text_input("URL for hyperlink:")
    for seg in segments:
        if seg["name"] in link_tables and link_col and link_url:
            tbl = seg["df"]
            if link_col in tbl.columns and 0 <= link_row < len(tbl):
                base = str(tbl.at[link_row, link_col])
                tbl.at[link_row, link_col] = (
                    f'{base} – '
                    f'<font color="blue"><a href="{link_url}">link</a></font>'
                )

    # Generate PDF
    title = st.text_input("Proposal Title", os.path.splitext(uploaded.name)[0])
    if st.button("Generate PDF"):
        for seg in segments:
            seg["df"] = calculate_and_insert_totals(seg["df"])
        pdf_bytes = make_pdf(segments, title)
        st.download_button(
            "📥 Download PDF", data=pdf_bytes, file_name=f"{title}.pdf", mime="application/pdf"
        )


if __name__ == "__main__":
    main()
