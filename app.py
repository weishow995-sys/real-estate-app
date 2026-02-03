import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, RGBColor
import datetime
import io
import time

# 基礎設定
st.set_page_config(page_title="房地產評估系統 (穩定版)", layout="centered")
st.title("🏠 房地產一鍵評估系統 (V12)")

# API KEY (已校正)
API_KEY = "AIzaSyDhxiL9d_cmWHmgQ9cms3xkj_f8piJdT8c"

uploaded_file = st.file_uploader("請選擇謄本檔案", type=["pdf", "png", "jpg", "jpeg"])

if uploaded_file:
    if st.button("🚀 開始分析 (請點擊一次後耐心等候)"):
        with st.spinner("系統連線中，請稍候..."):
            try:
                genai.configure(api_key=API_KEY)
                # 2026 年環境下最穩定的模型標籤
                model = genai.GenerativeModel('gemini-2.0-flash')
                
                prompt = "請解析此謄本。提取：姓名、完整身分證、持分、戶籍地址、設定額、登記日期。以30年2.15%利率試算殘值，並計算(行情80%價值-餘額)。"
                
                mime = "application/pdf" if uploaded_file.name.lower().endswith(".pdf") else uploaded_file.type
                response = model.generate_content([prompt, {"mime_type": mime, "data": uploaded_file.getvalue()}])
                
                # 製作 Word
                doc = Document()
                doc.add_heading('房地產評估報告書', 0)
                doc.add_paragraph(response.text)
                
                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)
                st.success("分析成功！")
                st.download_button(label="📥 下載 Word 報告", data=buf, file_name="評估報告.docx")
            except Exception as e:
                if "429" in str(e):
                    st.error("⚠️ 伺服器目前排隊人數過多。請『不要』重新整理，靜候 1 分鐘後再點一次按鈕即可。")
                else:
                    st.error(f"連線異常，請稍後再試：{e}")
