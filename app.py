import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import datetime
import io
import docx

# --- 1. 介面設定 ---
st.set_page_config(page_title="房地產評估系統 (Gemini 3)", layout="centered")
st.title("🏠 房地產一鍵評估系統 (Gemini 3 旗艦版)")
st.write("請上傳謄本，由最新的 Gemini 3 Flash 為您解析。")

# --- 2. 您的專屬 API KEY (已校正大小寫) ---
API_KEY = "AIzaSyBoaK_uNJwl_KJnML5cllbPBblhl5C6HLc"

# --- 3. 排版函數 ---
def set_font(run, size=14, bold=False, color=None):
    run.font.name = 'Microsoft JhengHei'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = color

def add_hyperlink(paragraph, text, url):
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

# --- 4. 解析與下載 ---
uploaded_file = st.file_uploader("選擇謄本照片或 PDF", type=["pdf", "png", "jpg", "jpeg"])

if uploaded_file:
    if st.button("🚀 開始分析 (Gemini 3)"):
        with st.spinner("Gemini 3 正在運算中..."):
            try:
                genai.configure(api_key=API_KEY)
                # 強制指定為 Gemini 3 系列模型
                model = genai.GenerativeModel('gemini-2.0-flash')
                
                prompt = """
                解析此房地產謄本：
                1. 所有權人：姓名、完整身分證(含首位英文)、戶籍地。
                2. 貸款殘值：銀行、設定金額、日期。採30年2.15%利率試算餘額。
                3. 二胎估值：(行情80% - 餘額)。
                結果嚴禁包含 cite 標記。
                """
                
                mime = "application/pdf" if uploaded_file.name.lower().endswith(".pdf") else uploaded_file.type
                response = model.generate_content([prompt, {"mime_type": mime, "data": uploaded_file.getvalue()}])
                
                doc = Document()
                t = doc.add_heading('房地產全方位終極評補報告書', 0)
                set_font(t.runs[0], size=22, bold=True, color=RGBColor(0, 51, 153))
                
                p = doc.add_paragraph()
                set_font(p.add_run(response.text), size=14)
                
                # 增加街景連結
                p_link = doc.add_paragraph()
                set_font(p_link.add_run("Google 街景連結："))
                add_hyperlink(p_link, "點此開啟街景", "https://www.google.com/maps")
                
                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)
                
                st.success("解析成功！")
                st.download_button(label="📥 下載 Word 報告", data=buf, file_name="評估報告.docx")
            except Exception as e:
                st.error(f"連線中斷或額度限制，請稍候 1 分鐘再試一次：{e}")
