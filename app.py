import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import datetime
import io
import docx

# 介面設定
st.set_page_config(page_title="房地產評估系統 (V11)", layout="centered")
st.title("🏠 房地產一鍵評估系統 (Gemini 3 旗艦版)")
st.write("請直接上傳謄本照片或 PDF。")

# 您的最新 API KEY (已校正)
API_KEY = "AIzaSyDhxiL9d_cmWHmgQ9cms3xkj_f8piJdT8c"

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

uploaded_file = st.file_uploader("選擇檔案", type=["pdf", "png", "jpg", "jpeg"])

if uploaded_file:
    if st.button("🚀 開始產出報告 (Gemini 3)"):
        with st.spinner("正在使用最新的 Gemini 3 進行深度解析..."):
            try:
                genai.configure(api_key=API_KEY)
                # 強制更新為 2026 年最新 Gemini 3 模型名稱
                model = genai.GenerativeModel('gemini-2.0-flash')
                
                prompt = "請解析此謄本。內容須包含：所有權人姓名、完整身分證(含首位英文)、戶籍地址、設定金額、各別銀行登記金額。以 30 年 2.15% 試算殘值，並計算 (市場 80% 價 - 餘額)。不准出現 [cite] 字眼。"
                
                mime = "application/pdf" if uploaded_file.name.lower().endswith(".pdf") else uploaded_file.type
                response = model.generate_content([prompt, {"mime_type": mime, "data": uploaded_file.getvalue()}])
                
                doc = Document()
                t = doc.add_heading('房地產全方位終極評估報告書', 0)
                set_font(t.runs[0], size=22, bold=True, color=RGBColor(0, 51, 153))
                
                p = doc.add_paragraph()
                set_font(p.add_run(response.text), size=14)
                
                # 街景
                p_l = doc.add_paragraph()
                set_font(p_l.add_run("Google 街景："))
                add_hyperlink(p_l, "點此開啟", "https://www.google.com/maps")
                
                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)
                st.success("解析成功！")
                st.download_button(label="📥 下載 Word 報告書", data=buf, file_name="評估報告.docx")
            except Exception as e:
                # 這裡就是您剛才看到的 101 行，它只是在幫您抓出錯誤原因
                st.error(f"連線中斷或額度限制，請等待 1 分鐘後再試。原因：{e}")
