# app.py

import io, os, re, tempfile
import pandas as pd, requests, streamlit as st
from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import inch
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

LOGO_URL = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
FONTS = {
    "Barlow":  "https://raw.githubusercontent.com/scooter7/ExcelBudgetBox/main/fonts/Barlow-Regular.ttf",
    "DMSerif": "https://raw.githubusercontent.com/scooter7/ExcelBudgetBox/main/fonts/DMSerifDisplay-Regular.ttf",
}

def register_fonts():
    for name, url in FONTS.items():
        r = requests.get(url)
        path = tempfile.NamedTemporaryFile(delete=False, suffix=".ttf").name
        open(path, "wb").write(r.content)
        pdfmetrics.registerFont(TTFont(name, path))
register_fonts()

st.set_page_config(page_title="Proposal → PDF", layout="wide")

def load_and_prepare_dataframe(uploaded_file):
    fn = uploaded_file.name.lower()
    df = pd.read_excel(uploaded_file, header=1) if fn.endswith((".xls",".xlsx")) else pd.read_csv(uploaded_file, header=1)
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
    return df.applymap(lambda v: v.replace("Est.", "Estimated") if isinstance(v, str) else v)

def calculate_and_insert_totals(df):
    mcol, icol = "Monthly Amount", "Item Total"
    impcol, ccol = "Estimated Impressions", "Estimated Conversions"
    itot = df.loc[df["Service"]=="Total", icol].iat[-1] if icol in df and "Total" in df["Service"].values else None
    impv = df.loc[df["Service"]=="Est. Conversions", impcol].iat[-1] if impcol in df and "Est. Conversions" in df["Service"].values else None
    monthly_sum = pd.to_numeric(df.get(mcol, []), errors="coerce").sum()
    conv_sum    = pd.to_numeric(df.get(ccol, []), errors="coerce").sum()
    drop = {"Total","Est. Conversions","Estimated Conversions","Est. Impressions","Estimated Impressions"}
    df = df.loc[~df["Service"].isin(drop)].reset_index(drop=True)
    total = {c: "" for c in df.columns}
    total["Service"] = "Total"
    total[mcol] = monthly_sum if mcol in total else ""
    total[icol] = itot if itot is not None else ""
    total[impcol] = impv if impv is not None else ""
    total[ccol] = conv_sum if ccol in total else ""
    return pd.concat([df, pd.DataFrame([total])], ignore_index=True)

def make_pdf(df, title, link_col, link_row, link_url):
    pdf_df = df.copy()
    for col in pdf_df.columns:
        if "date" in col.lower():
            pdf_df[col] = pd.to_datetime(pdf_df[col], errors="coerce").dt.strftime("%m/%d/%Y")
    pdf_df = pdf_df.fillna("")
    for col in ("Monthly Amount","Item Total"):
        if col in pdf_df:
            pdf_df[col] = pd.to_numeric(pdf_df[col], errors="coerce").fillna(0).map(lambda x: f"${x:,.0f}")
    # — hyperlink insertion as "… – link" only —
    if link_col in pdf_df.columns and link_row > 1 and link_url:
        idx = link_row - 2
        if 0 <= idx < len(pdf_df):
            base = str(pdf_df.at[idx, link_col])
            pdf_df.at[idx, link_col] = f'{base} – <link href="{link_url}">link</link>'
    # — build PDF —
    buf = io.BytesIO()
    tw, th = 17*inch, 11*inch
    doc = SimpleDocTemplate(buf, pagesize=(tw,th),
                            leftMargin=0.5*inch, rightMargin=0.5*inch,
                            topMargin=0.5*inch, bottomMargin=0.5*inch)
    elems=[]
    styles = getSampleStyleSheet()
    hdr = ParagraphStyle("hdr", parent=styles["BodyText"], fontName="DMSerif", fontSize=8, leading=9, alignment=1)
    bod = ParagraphStyle("bod", parent=styles["BodyText"], fontName="Barlow",  fontSize=7, leading=8, alignment=0)
    r = requests.get(LOGO_URL)
    img = PILImage.open(io.BytesIO(r.content))
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    img.save(tmp.name)
    elems.append(Image(tmp.name, width=4.5*inch, height=1.5*inch)); elems.append(Spacer(1,12))
    elems.append(Paragraph(f"<b>{title}</b>", styles["Title"])); elems.append(Spacer(1,12))
    header_cells=[Paragraph(c,hdr) for c in pdf_df.columns]
    data=[header_cells]
    for row in pdf_df.itertuples(False):
        cells=[]
        for c,v in zip(pdf_df.columns, row):
            t = "" if v is None else str(v)
            if "<link " in t or c in ("Service","Description","Notes"):
                cells.append(Paragraph(t,bod))
            else:
                cells.append(t)
        data.append(cells)
    widths=[0.12,0.30,0.06,0.08,0.08,0.12,0.12,0.06,0.06]
    colw=[doc.width*w for w in widths]
    table=Table(data, colWidths=colw, repeatRows=1)
    table.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.4,colors.black),
        ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
        ("FONTNAME",(0,0),(-1,0),"DMSerif"),
        ("FONTNAME",(0,1),(-1,-1),"Barlow"),
        ("FONTSIZE",(0,0),(-1,-1),7),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("BACKGROUND",(0,len(data)-1),(-1,len(data)-1),colors.lightgrey),
        ("FONTNAME",(0,len(data)-1),(-1,len(data)-1),"DMSerif"),
    ]))
    elems.append(table)
    doc.build(elems)
    buf.seek(0)
    return buf

def main():
    st.title("🔄 Proposal → PDF")
    up = st.file_uploader("Upload Excel/CSV", type=["xls","xlsx","csv"])
    if not up: return

    df = load_and_prepare_dataframe(up)
    st.dataframe(df.head())

    df = transform_service_column(df)
    df = replace_est(df)

    drop_cols = st.multiselect("Remove columns:", options=df.columns.tolist())
    if drop_cols:
        df = df.drop(columns=drop_cols, errors="ignore")

    link_col = st.selectbox("Hyperlink column:", [""]+df.columns.tolist())
    link_row = st.number_input("Hyperlink Excel row (1=header):", min_value=1, max_value=len(df)+1, value=2)
    link_url = st.text_input("URL for link:")

    title = st.text_input("Proposal Title", os.path.splitext(up.name)[0])

    if st.button("Generate PDF"):
        with st.spinner("Rendering PDF…"):
            df2 = calculate_and_insert_totals(df.copy())
            pdf = make_pdf(df2, title, link_col, link_row, link_url)
        st.download_button("📥 Download PDF", data=pdf, file_name=f"{title}.pdf", mime="application/pdf")

if __name__=="__main__":
    main()
