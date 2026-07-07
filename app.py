import streamlit as st

# 1. 網頁側邊欄：全科動態選擇
st.sidebar.header("🎯 考卷單元設定")
subject = st.sidebar.selectbox("選擇科目", ["國一生物", "國二理化", "國三理化", "地球科學"])

# 根據科目，動態顯示回數
if subject == "國一生物":
    exam_no = st.sidebar.selectbox("選擇回數", ["第1回_生命世界", "第2回_養分", "第3回_協調與恆定"])
elif subject == "國二理化":
    exam_no = st.sidebar.selectbox("選擇回數", ["第1回_基本測量", "第2回_物質的世界", "第3回_莫耳濃度"])

# 2. 自動在雲端抓取對應的 Markdown 標準答案檔
md_filename = f"answer_keys/{exam_no}.md"
try:
    with open(md_filename, "r", encoding="utf-8") as f:
        standard_answer_md = f.read()
    st.sidebar.success(f"✅ 已成功載入：{exam_no} 標準詳解庫")
except FileNotFoundError:
    st.sidebar.warning("⚠️ 尚未建立此單元的 Markdown 詳解庫，請先上傳。")

# 3. 背景自動讀取 API Key (完全免輸入！)
# 只要在 Streamlit 後台設定好，這行就能直接運作，畫面上不用給人填
api_key = st.secrets.get("GEMINI_API_KEY", "")
