import streamlit as st
import google.generativeai as genai
from docx import Document
import io

# 基礎頁面設定
st.set_page_config(page_title="房地產評估系統 (V13-穩定版)", layout="centered")
st.title("🏠 房地產一鍵評估系統 (V13)")
st.write("解析完畢後下載 Word，關閉分頁即刪除資料。")

# 您的 API KEY (已校正 K 為大寫)
API_KEY = "AIzaSyDhxiL9d_cmWHmgQ9cms3xkj_f8piJdT8c"

uploaded_file = st.file_uploader("選擇檔案", type=["pdf", "png", "jpg", "jpeg"])

if uploaded_file:
    # 增加提示，避免使用者連續點擊
    btn = st.button("🚀 啟動解析 (請點擊一次後靜候 15 秒)")
    if btn:
        with st.spinner("AI 正在連線中... 如果出現紅字請等 1 分鐘再試。"):
            try:
                genai.configure(api_key=API_KEY)
                # 使用額度最充足的 1.5-flash-8b 模型
                model = genai.GenerativeModel('gemini-1.5-flash-8b')
                
                prompt = "請解析此房地產謄本，提取：姓名、完整身分證(含首位英文)、地址、設定額、登記日期。試算目前餘額。結果禁止出現 [cite] 標記。"
                
                mime = "application/pdf" if uploaded_file.name.lower().endswith(".pdf") else uploaded_file.type
                response = model.generate_content([prompt, {"mime_type": mime, "data": uploaded_file.getvalue()}])
                
                # 生成 Word
                doc = Document()
                doc.add_heading('房地產評估報告', 0)
                doc.add_paragraph(response.text)
                
                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)
                st.success("🎉 解析成功！")
                st.download_button(label="📥 下載 Word 報告書", data=buf, file_name="評估報告.docx")
            except Exception as e:
                st.error(f"連線暫時忙碌，請等待 1 分鐘再按一次按鈕。原因：{e}")
