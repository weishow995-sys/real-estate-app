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
st.title("🏠 房地產一鍵評估系統 (V6)")
st.write("請上傳謄本照片或 PDF，系統將自動生成 Word 報告。")

# --- 2. 您的 API KEY (請在下方引號內貼上你的金鑰) ---
# ⚠️ 請確認這裡有換成你那串 AIza... 的代碼
API_KEY = "AIzaSyBoaK_uNJwI_KJnML5cllbPBbIhl5C6HLc"

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

# --- 4. 檔案上傳與按鈕 ---
uploaded_file = st.file_uploader("選擇檔案", type=["pdf", "png", "jpg", "jpeg"])

# 這裡改簡單了：只要有檔案，按鈕就出現
if uploaded_file:
    if st.button("🚀 點此開始產出報告"):
        if "您的_API_KEY" in API_KEY:
            st.error("錯誤：請先回到 GitHub 的第 19 行填入您的 API 金鑰！")
        else:
            with st.spinner("AI 正在深度解析中..."):
                try:
                    genai.configure(api_key=API_KEY)
                    model = genai.GenerativeModel('gemini-1.5-flash')
                    
                    prompt = """
                    請解析此房地產謄本，產出以下格式：
                    1. 所有權人：姓名、完整身分證(含首位英文)、持分、地址。
                    2. 貸款殘值：銀行名稱、設定額、日期。採30年2.15%利率計算餘額。
                    3. 二胎空間：計算(市場80%價格 - 剩餘貸款)。
                    嚴禁包含 標記。
                    """
                    
                    mime_type = "application/pdf" if uploaded_file.name.lower().endswith(".pdf") else uploaded_file.type
                    response = model.generate_content([prompt, {"mime_type": mime_type, "data": uploaded_file.getvalue()}])
                    
                    doc = Document()
                    title = doc.add_heading('', 0)
                    run_t = title.add_run('房地產全方位終極評估報告書')
                    set_font(run_t, size=20, bold=True, color=RGBColor(0, 51, 153))
                    
                    p = doc.add_paragraph()
                    set_font(p.add_run(response.text), size=14)
                    
                    # 增加街景連結
                    p_link = doc.add_paragraph()
                    set_font(p_link.add_run("Google 街景圖："))
                    add_hyperlink(p_link, "點此開啟街景", "https://www.google.com/maps")
                    
                    buf = io.BytesIO()
                    doc.save(buf)
                    buf.seek(0)
                    
                    st.success("評估完成！")
                    st.download_button(label="📥 下載 Word 報告書", data=buf, file_name="房產評估報告.docx")
                except Exception as e:
                    st.error(f"分析失敗，請確認 API Key 是否正確：{e}")
