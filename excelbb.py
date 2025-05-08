# app.py

import io, os, re, tempfile
import pandas as pd, requests, streamlit as st
from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import inch
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import (
    Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Logo and fonts
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
    open(path, "wb").write(r.content)
    pdfmetrics.registerFont(TTFont(name, path))

st.set_page_config(page_title="Proposal → PDF", layout="wide")


def load_and_prepare_dataframe(uploaded_file):
    """Load file with row 2 as header and rename column A to 'Service'."""
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
    """Clean the 'Service' column."""
    def clean(s):
        s = str(s)
        s = re.sub(r'^..:', '', s)       # remove leading XX:
        s = s.split('/', 1)[0]           # remove slash+suffix
        s = re.sub(r'\(.*?\)', '', s)  # remove parentheses
        return s.strip()
    df["Service"] = df["Service"].fillna("").apply(clean)
    return df


def replace_est(df):
    """Rename Est.→Estimated in column names and values."""
    df.columns = [c.replace("Est.", "Estimated").strip() for c in df.columns]
    return df.applymap(lambda v: v.replace("Est.", "Estimated") if isinstance(v, str) else v)


def split_tables(df):
    """Split the DataFrame into segments ending with rows where Service=='Total'."""
    segments = []
    start = 0
    for i, val in enumerate(df['Service']):
        if str(val).strip().lower() == 'total':
            segments.append(df.iloc[start:i+1].reset_index(drop=True))
            start = i+1
    # capture trailing rows without a total
    if start < len(df):
        segments.append(df.iloc[start:].reset_index(drop=True))
    return segments


def calculate_and_insert_totals(df_segment):
    """Recalculate totals for a segment and append a Total row."""
    mcol, icol = "Monthly Amount", "Item Total"
    impcol, ccol = "Estimated Impressions", "Estimated Conversions"

    # sum numeric
    monthly_sum = pd.to_numeric(df_segment.get(mcol, []), errors="coerce").sum()
    conv_sum = pd.to_numeric(df_segment.get(ccol, []), errors="coerce").sum()

    # build new total row
    total = {c: "" for c in df_segment.columns}
    total['Service'] = 'Total'
    if mcol in total: total[mcol] = monthly_sum
    if ccol in total: total[ccol] = conv_sum

    return pd.concat([df_segment, pd.DataFrame([total])], ignore_index=True)


def make_pdf(segments, title):
    """Render multiple table segments to a single PDF."""
    buf = io.BytesIO()
    tw, th = 17 * inch, 11 * inch
    doc = SimpleDocTemplate(
        buf, pagesize=(tw, th),
        leftMargin=0.5*inch, rightMargin=0.5*inch,
        topMargin=0.5*inch, bottomMargin=0.5*inch,
    )

    styles = getSampleStyleSheet()
    hdr_style = ParagraphStyle('hdr', parent=styles['BodyText'], fontName='DMSerif', fontSize=8, leading=9, alignment=1)
    body_style = ParagraphStyle('bod', parent=styles['BodyText'], fontName='Barlow', fontSize=7, leading=8, alignment=0)

    elems = []
    # logo
    resp = requests.get(LOGO_URL)
    logo_img = PILImage.open(io.BytesIO(resp.content))
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
    logo_img.save(tmp.name)
    elems.append(Image(tmp.name, width=4.5*inch, height=1.5*inch))
    elems.append(Spacer(1, 12))
    elems.append(Paragraph(f"<b>{title}</b>", styles['Title']))
    elems.append(Spacer(1, 12))

    for idx, seg in enumerate(segments):
        # recalc totals
        seg_final = calculate_and_insert_totals(seg)

        # format dates & currency
        for col in seg_final.columns:
            if 'date' in col.lower():
                seg_final[col] = pd.to_datetime(seg_final[col], errors='coerce').dt.strftime('%m/%d/%Y')
        for col in ('Monthly Amount','Item Total'):
            if col in seg_final.columns:
                seg_final[col] = pd.to_numeric(seg_final[col], errors='coerce').fillna(0).map(lambda x: f"${x:,.0f}")

        # build table data
        header = [Paragraph(c, hdr_style) for c in seg_final.columns]
        data = [header]
        for row in seg_final.itertuples(index=False):
            cells = []
            for c, v in zip(seg_final.columns, row):
                t = '' if v is None else str(v)
                cells.append(Paragraph(t, body_style) if c in ('Service','Description','Notes') or '<link' in t else t)
            data.append(cells)

        col_widths = [doc.width * w for w in [0.12,0.30,0.06,0.08,0.08,0.12,0.12,0.06,0.06]]
        table = Table(data, colWidths=col_widths, repeatRows=1)
        table.setStyle(TableStyle([
            ('GRID',(0,0),(-1,-1),0.4,colors.black),
            ('BACKGROUND',(0,0),(-1,0),colors.lightgrey),
            ('FONTNAME',(0,0),(-1,0),'DMSerif'),
            ('FONTNAME',(0,1),(-1,-1),'Barlow'),
            ('FONTSIZE',(0,0),(-1,-1),7),
            ('ALIGN',(0,0),(-1,-1),'CENTER'),
            ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
            ('BACKGROUND',(0,len(data)-1),(-1,len(data)-1),colors.lightgrey),
            ('FONTNAME',(0,len(data)-1),(-1,len(data)-1),'DMSerif'),
        ]))

        elems.append(table)
        elems.append(Spacer(1, 24))

    doc.build(elems)
    buf.seek(0)
    return buf


def main():
    st.title("🔄 Proposal → PDF")

    uploaded = st.file_uploader("Upload Excel/CSV", type=["xls","xlsx","csv"])
    if not uploaded:
        return

    # 1) load & prepare
    df = load_and_prepare_dataframe(uploaded)
    df = transform_service_column(df)
    df = replace_est(df)

    # 2) split into tables
    segments = split_tables(df)

    # 3) allow selecting tables to modify
    table_choices = list(range(len(segments)))
    chosen = st.multiselect(
        "Select tables to modify (by number):", 
        options=table_choices,
        format_func=lambda i: f"Table {i+1}"
    )

    # 4) drop columns for chosen tables
    drop_cols = st.multiselect("Columns to remove:", options=df.columns.tolist())
    for i in chosen:
        segments[i] = segments[i].drop(columns=drop_cols, errors='ignore')

    # 5) hyperlink settings for chosen tables
    link_tables = st.multiselect(
        "Tables to add hyperlink to:", 
        options=chosen, 
        format_func=lambda i: f"Table {i+1}"
    )
    link_col = st.selectbox("Hyperlink column:", options=[""]+df.columns.tolist())
    link_row = st.number_input("Excel row in table to hyperlink (1=first data row):", min_value=1, value=1)
    link_url = st.text_input("URL for link:")
    for i in link_tables:
        # apply hyperlink per segment
        seg = segments[i]
        idx = link_row - 1
        if link_col in seg.columns and 0 <= idx < len(seg) and link_url:
            base = str(seg.at[idx, link_col])
            segments[i].at[idx, link_col] = f'{base} – <link href="{link_url}">link</link>'

    # 6) preview each segment
    for idx, seg in enumerate(segments):
        st.subheader(f"Table {idx+1}")
        st.dataframe(seg.reset_index(drop=True))

    # 7) title and PDF
    proposal_title = st.text_input("Proposal Title", os.path.splitext(uploaded.name)[0])
    if st.button("Generate PDF"):
        with st.spinner("Rendering PDF…"):
            pdf_bytes = make_pdf(segments, proposal_title)
        st.download_button("📥 Download PDF", data=pdf_bytes, file_name=f"{proposal_title}.pdf", mime="application/pdf")

if __name__ == '__main__':
    main()
