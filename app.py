import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import datetime
import io
import docx

# 網頁基礎設定
st.set_page_config(page_title="房地產終極評估系統", layout="centered")
st.title("🏠 房地產一鍵評估系統")
st.write("上傳謄本照片，直接生成 Word 評估報告。")

# 請將您的 API Key 填入下方引號中
API_KEY = "AIzaSyBoaK_uNJwI_KJnML5cllbPBbIhl5C6HLc"

def set_font(run, size=14, bold=False, color=None):
    run.font.name = 'Microsoft JhengHei'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = color

def add_hyperlink(paragraph, url, text):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(docx.oxml.shared.qn('r:id'), r_id)
    new_run = docx.oxml.shared.OxmlElement('w:r')
    rPr = docx.oxml.shared.OxmlElement('w:rPr')
    c = docx.oxml.shared.OxmlElement('w:color')
    c.set(docx.oxml.shared.qn('w:val'), '0000FF')
    rPr.append(c)
    u = docx.oxml.shared.OxmlElement('w:u')
    u.set(docx.oxml.shared.qn('w:val'), 'single')
    rPr.append(u)
    f = docx.oxml.shared.OxmlElement('w:rFonts')
    f.set(docx.oxml.shared.qn('w:eastAsia'), 'Microsoft JhengHei')
    rPr.append(f)
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

def calc_balance(principal, rate, years, months):
    r = rate/100/12
    n = years*12
    if r == 0: return principal * (1 - months/n)
    return principal * ((1+r)**n - (1+r)**months) / ((1+r)**n - 1)

uploaded_file = st.file_uploader("請選擇謄本檔案 (PDF/JPG/PNG)", type=["pdf", "png", "jpg", "jpeg"])

if uploaded_file and API_KEY != "您的_API_KEY_貼在這邊":
    if st.button("🚀 開始分析"):
        with st.spinner("系統分析中..."):
            genai.configure(api_key=API_KEY)
            model = genai.GenerativeModel('gemini-1.5-pro')
            prompt = "請深度解析此謄本。包含產權警示、RC/SRC建材、屋齡、姓名、完整身分證(含英文字母)、持分、戶籍地。計算30年利率2.15%殘值、市場80%價格與二胎估值。禁止cite標記。"
            res = model.generate_content([prompt, {"mime_type": uploaded_file.type, "data": uploaded_file.getvalue()}])
            
            doc = Document()
            title = doc.add_heading('', 0)
            run_t = title.add_run('房地產全方位終極評估報告書')
            set_font(run_t, size=22, bold=True, color=RGBColor(0, 51, 153))
            doc.add_paragraph(res.text) # 簡化寫入，實際會按表格排版
            
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            st.success("評估完成")
            st.download_button("📥 下載 Word 報告", data=buf, file_name="房產評估.docx")
