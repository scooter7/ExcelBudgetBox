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
    """Load file with row 2 as header and rename col A to 'Service'."""
    fn = uploaded_file.name.lower()
    if fn.endswith((".xls", ".xlsx")):
        df = pd.read_excel(uploaded_file, header=1)
    else:
        df = pd.read_csv(uploaded_file, header=1)
    # rename first column to Service
    first = df.columns[0]
    if first != "Service":
        df = df.rename(columns={first: "Service"})
    return df


def transform_service_column(df):
    """Clean the 'Service' column entries."""
    def clean(s):
        s = str(s)
        s = re.sub(r'^..:', '', s)       # remove leading XX:
        s = s.split('/', 1)[0]           # drop slash and after
        s = re.sub(r'\(.*?\)', '', s)  # remove parentheses content
        return s.strip()
    df['Service'] = df['Service'].fillna('').apply(clean)
    return df


def replace_est(df):
    """Rename Est.→Estimated and strip column names."""
    df.columns = [c.replace('Est.', 'Estimated').strip() for c in df.columns]
    return df.applymap(lambda v: v.replace('Est.', 'Estimated') if isinstance(v, str) else v)


def split_tables(df):
    """Split into segments ending with a 'Total' row, label by first Service value."""
    segments = []
    start = 0
    for i, svc in enumerate(df['Service']):
        if svc.strip().lower() == 'total':
            seg = df.iloc[start:i+1].reset_index(drop=True)
            # derive name: first non-empty Service that's not header
            name = next((x for x in seg['Service'] if x and x.strip().lower() not in ['service']), 'Table')
            segments.append({'name': name, 'df': seg})
            start = i+1
    # handle trailing rows without Total
    if start < len(df):
        seg = df.iloc[start:].reset_index(drop=True)
        name = next((x for x in seg['Service'] if x and x.strip().lower() not in ['service']), 'Table')
        segments.append({'name': name, 'df': seg})
    return segments


def calculate_and_insert_totals(seg_df):
    """Recalculate segment totals and append."""
    df = seg_df.copy()
    mcol, icol = 'Monthly Amount', 'Item Total'
    ccol = 'Estimated Conversions'
    # sums
    monthly_sum = pd.to_numeric(df.get(mcol, []), errors='coerce').sum()
    conv_sum = pd.to_numeric(df.get(ccol, []), errors='coerce').sum()
    # drop legacy footers
    drop = ['total', 'est. conversions', 'estimated conversions']
    df = df.loc[~df['Service'].str.strip().str.lower().isin(drop)].reset_index(drop=True)
    # append new total row
    total = {c: '' for c in df.columns}
    total['Service'] = 'Total'
    if mcol in total: total[mcol] = monthly_sum
    if ccol in total: total[ccol] = conv_sum
    return pd.concat([df, pd.DataFrame([total])], ignore_index=True)


def make_pdf(segments, title):
    """Build a tabloid PDF with each named segment."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=(17*inch, 11*inch),
        leftMargin=0.5*inch, rightMargin=0.5*inch,
        topMargin=0.5*inch, bottomMargin=0.5*inch
    )
    styles = getSampleStyleSheet()
    hdr = ParagraphStyle('hdr', parent=styles['BodyText'], fontName='DMSerif', fontSize=8, leading=9, alignment=1)
    bod = ParagraphStyle('bod', parent=styles['BodyText'], fontName='Barlow', fontSize=7, leading=8, alignment=0)

    elems = []
    # logo
    r = requests.get(LOGO_URL)
    logo_img = PILImage.open(io.BytesIO(r.content))
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
    logo_img.save(tmp.name)
    elems.append(Image(tmp.name, width=4.5*inch, height=1.5*inch))
    elems.append(Spacer(1,12))
    elems.append(Paragraph(f'<b>{title}</b>', styles['Title']))
    elems.append(Spacer(1,12))

    for seg in segments:
        name = seg['name']
        df = seg['df']
        # clean empty/service header rows
        df = df[df['Service'].notna() & df['Service'].str.strip().astype(bool)]
        df = df.loc[~((df['Service'].str.strip().str.lower()=='service') &
                      (df.get('Description','').str.strip().str.lower()=='description'))]
        # recalc totals
        df = calculate_and_insert_totals(df)
        # format dates & currency
        for col in df.columns:
            if 'date' in col.lower():
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%m/%d/%Y')
        for col in ('Monthly Amount','Item Total'):
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).map(lambda x: f'${x:,.0f}')
        # segment title
        elems.append(Paragraph(f'<b>{name}</b>', styles['Heading2']))
        elems.append(Spacer(1,6))
        # build table
        hdr_cells = [Paragraph(c, hdr) for c in df.columns]
        data = [hdr_cells]
        for row in df.itertuples(index=False):
            cells = []
            for c, v in zip(df.columns, row):
                txt = '' if v is None else str(v)
                if '<a href' in txt or c in ('Service','Description','Notes'):
                    cells.append(Paragraph(txt, bod))
                else:
                    cells.append(txt)
            data.append(cells)
        col_widths = [doc.width*w for w in [0.12,0.30,0.06,0.08,0.08,0.12,0.12,0.06,0.06]]
        tbl = Table(data, colWidths=col_widths, repeatRows=1)
        tbl.setStyle(TableStyle([
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
        elems.append(tbl)
        elems.append(Spacer(1,24))

    doc.build(elems)
    buf.seek(0)
    return buf


def main():
    st.title('🔄 Proposal → PDF')
    up = st.file_uploader('Upload Excel/CSV', type=['xls','xlsx','csv'])
    if not up:
        return

    # load, clean, split
    df = load_and_prepare_dataframe(up)
    df = transform_service_column(df)
    df = replace_est(df)
    segments = split_tables(df)

    # choose which segments to modify
    names = [seg['name'] for seg in segments]
    choose = st.multiselect('Modify tables:', options=names)

    # drop columns for chosen
    drop_cols = st.multiselect('Remove columns:', options=df.columns.tolist())
    for seg in segments:
        if seg['name'] in choose and drop_cols:
            seg['df'] = seg['df'].drop(columns=drop_cols, errors='ignore')

    # hyperlink for chosen
    hl_choose = st.multiselect('Add hyperlink to tables:', options=names)
    link_col = st.selectbox('Hyperlink column:', options=['']+df.columns.tolist())
    link_row = st.number_input('Row in table to link (1=first data row):', min_value=1, value=1)
    link_url = st.text_input('URL for hyperlink:')
    for seg in segments:
        if seg['name'] in hl_choose and link_col and link_url:
            df_seg = seg['df']
            idx = link_row - 1
            if link_col in df_seg.columns and 0 <= idx < len(df_seg):
                base = str(df_seg.at[idx, link_col])
                df_seg.at[idx, link_col] = f'{base} – <link href="{link_url}">link</link>'

    # preview
    for seg in segments:
        st.subheader(seg['name'])
        st.dataframe(seg['df'].reset_index(drop=True))

    # title & PDF
    title = st.text_input('Proposal Title', os.path.splitext(up.name)[0])
    if st.button('Generate PDF'):
        pdf = make_pdf(segments, title)
        st.download_button('📥 Download PDF', data=pdf, file_name=f'{title}.pdf', mime='application/pdf')

if __name__ == '__main__':
    main()
