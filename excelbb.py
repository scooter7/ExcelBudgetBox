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
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Carnegie logo
LOGO_URL = (
    "https://www.carnegiehighered.com/wp-content/uploads/2021/11/"
    "Twitter-Image-2-2021.png"
)
# Font URLs
FONTS = {
    "Barlow":  "https://raw.githubusercontent.com/scooter7/ExcelBudgetBox/main/fonts/Barlow-Regular.ttf",
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
    # Load with row 2 as header, rename col A to Service
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
    df.columns = [c.replace("Est.", "Estimated") for c in df.columns]
    df.columns = [c.strip() for c in df.columns]
    return df.applymap(lambda v: v.replace("Est.", "Estimated") if isinstance(v, str) else v)


def drop_columns(df):
    drop = st.multiselect("Remove columns:", options=df.columns.tolist())
    if drop:
        df = df.drop(columns=drop, errors="ignore")
    return df, drop


def insert_hyperlink(df, link_col, link_row, link_url):
    if link_col in df.columns and link_row > 1 and link_url:
        idx = link_row - 2
        if 0 <= idx < len(df):
            base = str(df.at[idx, link_col])
            df.at[idx, link_col] = f'{base} – <link href="{link_url}">link</link>'
    return df


def split_tables(df):
    # Segment by legacy Total rows
    segments = []
    start = 0
    for i, val in enumerate(df['Service']):
        if str(val).strip().lower() == 'total':
            seg = df.iloc[start:i+1].reset_index(drop=True)
            segments.append(seg)
            start = i+1
    return segments


def calculate_and_insert_totals(df):
    mcol, icol = "Monthly Amount", "Item Total"
    impcol, ccol = "Estimated Impressions", "Estimated Conversions"
    itot = None
    if icol in df.columns and 'Total' in df['Service'].values:
        itot = df.loc[df['Service']=='Total', icol].iat[-1]
    impv = None
    if impcol in df.columns and 'Est. Conversions' in df['Service'].values:
        impv = df.loc[df['Service']=='Est. Conversions', impcol].iat[-1]
    ms = pd.to_numeric(df.get(mcol, []), errors='coerce').sum()
    cs = pd.to_numeric(df.get(ccol, []), errors='coerce').sum()
    drop_set = {'Total','Est. Conversions','Estimated Conversions','Est. Impressions','Estimated Impressions'}
    df_clean = df.loc[~df['Service'].isin(drop_set)].reset_index(drop=True)
    total = {c: '' for c in df_clean.columns}
    total['Service'] = 'Total'
    total[mcol] = ms
    if itot is not None: total[icol] = itot
    if impv is not None: total[impcol] = impv
    total[ccol] = cs
    return pd.concat([df_clean, pd.DataFrame([total])], ignore_index=True)


def make_pdf(segments, title):
    buf = io.BytesIO()
    tw, th = 17*inch, 11*inch
    doc = SimpleDocTemplate(buf, pagesize=(tw,th), leftMargin=0.5*inch,
                            rightMargin=0.5*inch, topMargin=0.5*inch, bottomMargin=0.5*inch)
    styles = getSampleStyleSheet()
    hdr = ParagraphStyle('hdr', parent=styles['BodyText'], fontName='DMSerif', fontSize=8, leading=9, alignment=1)
    bod = ParagraphStyle('bod', parent=styles['BodyText'], fontName='Barlow', fontSize=7, leading=8, alignment=0)
    elems = []
    # logo
    r = requests.get(LOGO_URL)
    img = PILImage.open(io.BytesIO(r.content))
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
    img.save(tmp.name)
    elems.append(Image(tmp.name, width=4.5*inch, height=1.5*inch))
    elems.append(Spacer(1,12))
    elems.append(Paragraph(f"<b>{title}</b>", styles['Title']))
    elems.append(Spacer(1,12))
    # each table
    for seg in segments:
        seg2 = calculate_and_insert_totals(seg)
        # drop blank/repeated headers
        seg2 = seg2.loc[~((seg2['Service']=='Service') & (seg2.get('Description','')=='Description'))]
        # build data
        header_cells = [Paragraph(c, hdr) for c in seg2.columns]
        data = [header_cells]
        for row in seg2.itertuples(index=False):
            cells = []
            for c,v in zip(seg2.columns, row):
                txt = '' if v is None else str(v)
                if '<link ' in txt or c in ('Service','Description','Notes'):
                    cells.append(Paragraph(txt, bod))
                else:
                    cells.append(txt)
            data.append(cells)
        widths = [0.12,0.30,0.06,0.08,0.08,0.12,0.12,0.06,0.06]
        colw = [doc.width*w for w in widths]
        table = Table(data, colWidths=colw, repeatRows=1)
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
        elems.append(Spacer(1,24))
    doc.build(elems)
    buf.seek(0)
    return buf


def main():
    st.title('🔄 Proposal → PDF')
    up = st.file_uploader('Upload Excel/CSV', type=['xls','xlsx','csv'])
    if not up: return
    df = load_and_prepare_dataframe(up)
    df = transform_service_column(df)
    df = replace_est(df)
    df, drop_cols = (lambda d: (d.drop(columns=st.multiselect('Remove columns:', options=d.columns.tolist()), errors='ignore'), None))(df)
    st.dataframe(df.head())
    link_col = st.selectbox('Hyperlink column:', ['']+df.columns.tolist())
    link_row = st.number_input('Hyperlink Excel row (1=header):', min_value=1, max_value=len(df)+1, value=2)
    link_url = st.text_input('URL for link:')
    if link_col:
        df = insert_hyperlink(df, link_col, link_row, link_url)
    title = st.text_input('Proposal Title', os.path.splitext(up.name)[0])
    if st.button('Generate PDF'):
        with st.spinner('Rendering PDF…'):
            segments = split_tables(df)
            pdf = make_pdf(segments, title)
        st.download_button('📥 Download PDF', data=pdf, file_name=f'{title}.pdf', mime='application/pdf')

if __name__ == '__main__':
    main()
