import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import datetime
import io
import docx

# --- 1. 網頁介面設定 ---
st.set_page_config(page_title="房地產評估系統", layout="centered")
st.title("🏠 房地產一鍵評估系統 (V7-穩定版)")
st.write("上傳謄本照片或 PDF，即可自動生成 Word 報告。")

# --- 2. API KEY 設定 ---
# ⚠️ 請確保下方引號內是您在 AI Studio 申請的 AIzaSy... 代碼
API_KEY = "AIzaSyBoak_uNJwl_KJnML5cllbPBblhl5C6HLc"

# --- 3. 工具函數 ---
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

# --- 4. 檔案上傳與處理 ---
uploaded_file = st.file_uploader("請選擇謄本檔案", type=["pdf", "png", "jpg", "jpeg"])

if uploaded_file:
    if st.button("🚀 點此產出評估報告"):
        with st.spinner("AI 正在解析數據中..."):
            try:
                genai.configure(api_key=API_KEY)
                # 更新模型名稱至 2026 穩定版
                model = genai.GenerativeModel('gemini-2.0-flash')
                
                prompt = """
                請解析此房地產謄本，產出以下格式：
                1. 所有權人：姓名、完整身分證(含首位英文)、持分比例、戶籍地址。
                2. 貸款殘值：銀行名稱、設定額、登記日期。
                3. 二胎空間試算：以設定金額除以 1.2 作為本金，採 30 年 2.15% 利率試算目前餘額。
                4. 計算：(市場行情 80% 價值 - 目前餘額)。
                結果嚴禁包含任何 標記。
                """
                
                mime_type = "application/pdf" if uploaded_file.name.lower().endswith(".pdf") else uploaded_file.type
                response = model.generate_content([prompt, {"mime_type": mime_type, "data": uploaded_file.getvalue()}])
                
                # --- 製作 Word 檔案 ---
                doc = Document()
                title = doc.add_heading('', 0)
                run_t = title.add_run('房地產全方位終極評估報告書')
                set_font(run_t, size=20, bold=True, color=RGBColor(0, 51, 153))
                
                # 報告正文
                p = doc.add_paragraph()
                set_font(p.add_run(response.text), size=14)
                
                # 增加街景連結區塊
                doc.add_heading('', level=1).add_run('外部資源連結').font.size = Pt(16)
                p_link = doc.add_paragraph()
                set_font(p_link.add_run("Google 街景服務："))
                add_hyperlink(p_link, "點此開啟街景圖", "https://www.google.com/maps")
                
                # 轉存為二進位流供下載
                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)
                
                st.success("評估完成！")
                st.download_button(label="📥 下載 Word 報告書", data=buf, file_name=f"評估報告_{datetime.date.today()}.docx")
                
            except Exception as e:
                st.error(f"分析失敗，建議更換模型名稱：{e}")
