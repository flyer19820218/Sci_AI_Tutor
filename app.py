# =====================================================================
# 🧬 生物科個人化診斷報告產生器 v2.0（Streamlit 版）
# 功能：上傳 Excel → 自動產生個人化診斷報告 + 補考卷 → 下載 / 寄送
# =====================================================================

import streamlit as st
import pandas as pd
import os
import zipfile
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether, PageBreak
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import tempfile
from pathlib import Path

st.set_page_config(page_title="生物科診斷報告產生器", page_icon="🧬", layout="wide")

# =============================================
# 側邊欄：使用教學
# =============================================

with st.sidebar:
    st.markdown("## 📖 使用教學")
    st.markdown("---")

    with st.expander("📌 第一步：準備資料", expanded=True):
        st.markdown("""
        1. 下載 Excel 範本檔案（或自行建立）
        2. 填入學生資料，必須包含以下欄位：

        | 欄位名稱 | 說明 | 範例 |
        |----------|------|------|
        | 座號 | 學生座號 | 28 |
        | 姓名 | 學生姓名 | 廖禹晴 |
        | 總分 | 測驗得分 | 77.5 |
        | 基礎錯題 | 基礎題錯題編號 | 4,13,14,16,17 |
        | 精熟錯題 | 精熟題錯題編號 | 1,8 |
        | 分析文字 | 考點拆解（可用HTML） | 請參考下方說明 |
        | 思維診斷 | 學習認知診斷 | 請參考下方說明 |
        | 暖心引導 | 給學生的鼓勵話語 | 請參考下方說明 |
        | Email | 學生電子郵件 | student@email.com |
        """)

    with st.expander("✍️ 第二步：填入診斷文字"):
        st.markdown("""
        **分析文字**：可以用 HTML 標籤，例如：

        <b>1. 基礎題4【養分與熱量迷思】：</b>
        纖維素無法提供熱量...

        <b>2. 基礎題13【生物攝食】：</b>
        變形蟲用吞噬作用...

        **思維診斷**：描述學生的思考模式，例如：
        【邏輯大師，細節迷路】學生在精熟題全對，
        但基礎題的生理名詞記憶不夠精確...

        **暖心引導**：給學生鼓勵與挑戰，例如：
        昱O啊，先用力給自己的大腦鼓鼓掌！
        請你思考這三個問題：(1)...
        """)

    with st.expander("📤 第三步：上傳與產生"):
        st.markdown("""
        1. 點擊「瀏覽檔案」上傳你的 Excel
        2. 檢查資料預覽是否正確
        3. 點擊「產生所有學生的診斷報告」
        4. 等待系統生成（約 10~30 秒）
        5. 點擊「下載所有診斷報告」儲存 ZIP 檔
        """)

    with st.expander("📧 第四步：寄送 Email（選用）"):
        st.markdown("""
        **使用 Gmail 寄信前，請先設定「應用程式密碼」：**

        1. 開啟 Google 帳號的「兩步驟驗證」
        2. 到「應用程式密碼」產生 16 位密碼
        3. 在下方輸入：
           - 寄件者：你的 Gmail
           - 密碼：16 位應用程式密碼

        [點此前往 Google 應用程式密碼設定](https://myaccount.google.com/apppasswords)
        """)

    st.markdown("---")
    st.caption("💡 有問題？請洽詢生物科教學研究會")

# =============================================
# 主畫面
# =============================================

st.title("🧬 生物科個人化診斷報告產生器")
st.markdown("---")

col1, col2, col3 = st.columns(3)
with col1:
    font_exists = Path("NotoSansTC-Regular.ttf").exists()
    st.metric("📋 已載入字型", "✅ 就緒" if font_exists else "❌ 未找到")
with col2:
    if "df" in st.session_state:
        st.metric("📊 學生資料", f"{len(st.session_state.df)} 人")
    else:
        st.metric("📊 學生資料", "等待上傳")
with col3:
    if "zip_path" in st.session_state:
        st.metric("📄 報告狀態", "✅ 已產生")
    else:
        st.metric("📄 報告狀態", "未產生")

# =============================================
# 自動讀取字型檔
# =============================================

@st.cache_resource
def load_font():
    font_path = Path("NotoSansTC-Regular.ttf")
    if font_path.exists():
        pdfmetrics.registerFont(TTFont("CustomFont", str(font_path)))
        return True
    return False

if not load_font():
    st.error("""
    ❌ **找不到字型檔！**

    請確認 `NotoSansTC-Regular.ttf` 已放在專案根目錄。

    如果你還沒有這個檔案，可以從以下管道取得：
    - [Google Noto Sans TC 下載](https://fonts.google.com/noto/specimen/Noto+Sans+TC)
    - 或使用系統內建的其他中文字型（需調整程式碼）
    """)
    st.stop()

# =============================================
# 步驟一：上傳 Excel
# =============================================

st.markdown("## 📤 步驟一：上傳學生資料 Excel")

uploaded_file = st.file_uploader(
    "請上傳 Excel 檔案（.xlsx 格式）",
    type=["xlsx"],
    help="欄位名稱必須包含：座號、姓名、總分、基礎錯題、精熟錯題、分析文字、思維診斷、暖心引導、Email"
)

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.session_state.df = df
    st.success(f"✅ 成功讀取 {len(df)} 位學生資料！")

    required_cols = ["座號", "姓名", "總分", "基礎錯題", "精熟錯題"]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        st.warning(f"⚠️ 缺少以下欄位：{', '.join(missing_cols)}，請確認 Excel 格式是否正確")

    with st.expander("📋 點擊展開資料預覽"):
        st.dataframe(df, use_container_width=True)
        st.caption(f"共 {len(df)} 位學生，{len(df.columns)} 個欄位")

    if "分析文字" in df.columns:
        empty_analysis = df["分析文字"].isna().sum()
        if empty_analysis > 0:
            st.info(f"ℹ️ 有 {empty_analysis} 位學生的「分析文字」欄位為空白，將顯示預設文字")

# =============================================
# 步驟二：產生 PDF
# =============================================

st.markdown("## 🚀 步驟二：產生診斷報告")

with st.expander("⚙️ 進階選項"):
    show_class = st.checkbox("在報告中顯示班級名稱", value=True)
    class_name = st.text_input("班級名稱", value="9年01班")
    min_questions = st.number_input("每份報告最少題數", min_value=3, max_value=10, value=5)

# =============================================
# 錯題題庫
# =============================================

basic_question_map = {
    3: "科學探究的正確順序為：觀察 → ？ → 提出假設 → 實驗 (A)提出問題 (B)設計實驗 (C)得出結論 (D)修正假設",
    4: "人體攝取纖維素（膳食纖維）的主要目的為何？ (A)提供熱量 (B)刺激腸道蠕動 (C)合成蛋白質 (D)儲存脂肪",
    5: "植物表皮細胞的主要功能為何？ (A)光合作用 (B)保護作用 (C)運輸水分 (D)儲存養分",
    7: "製作玻片標本時，蓋玻片應如何操作才不易產生氣泡？ (A)直接平放壓下 (B)傾斜45°緩慢蓋下 (C)先滴一滴水再蓋 (D)用鑷子夾住快速放下",
    11: "使用複式顯微鏡時，將光圈由小調大，視野亮度會如何變化？ (A)變亮 (B)變暗 (C)不變 (D)視野範圍變大",
    13: "變形蟲攝取食物的方式為何？ (A)用牙齒咀嚼 (B)用管狀口器吸食 (C)用細胞膜進行吞噬作用 (D)用觸手捕捉",
    14: "捕蟲植物（如豬籠草）在大自然中的主要營養方式為何？ (A)完全依賴捕食昆蟲 (B)主要行光合作用，捕蟲只為補充氮元素 (C)只靠土壤中的養分 (D)屬於異營生物",
    15: "動物細胞中，DNA主要存在於哪一個構造中？ (A)細胞膜 (B)細胞質 (C)細胞核 (D)粒線體",
    16: "關於維生素的敘述，下列何者正確？ (A)可提供熱量 (B)人體可自行合成所有維生素 (C)缺乏時會影響生理機能 (D)主要來源為油脂",
    17: "人體中製造膽汁的器官為何？ (A)膽囊 (B)肝臟 (C)胰臟 (D)小腸",
}

advanced_question_map = {
    1: "關於大腸的敘述，下列何者正確？ (A)可分泌消化液 (B)主要吸收水分與形成糞便 (C)可分解蛋白質 (D)可吸收葡萄糖",
    2: "人體消化管中，哪一個部位的環境為強酸性？ (A)口腔 (B)胃 (C)小腸 (D)大腸",
    3: "某酵素在40°C時活性最高，在60°C時活性明顯下降。40°C稱為該酵素的什麼？ (A)最適溫度 (B)變性溫度 (C)起始溫度 (D)容忍溫度",
    4: "植物行光合作用所產生的氣體是？ (A)二氧化碳 (B)氧氣 (C)氮氣 (D)水蒸氣",
    5: "三大養分在人體消化道中被分解的順序（由先到後）為何？ (A)脂質→蛋白質→澱粉 (B)澱粉→蛋白質→脂質 (C)蛋白質→澱粉→脂質 (D)脂質→澱粉→蛋白質",
    8: "胃液中的酵素最適合在哪一種環境中作用？ (A)強酸性 (B)中性 (C)弱鹼性 (D)強鹼性",
    11: "藍綠菌（藍細菌）能行光合作用，但下列敘述何者正確？ (A)具有葉綠體 (B)屬於原核生物，沒有葉綠體 (C)屬於真核生物 (D)具有細胞核",
}


def generate_practice_questions(wrong_basic_str, wrong_advanced_str, min_q=5):
    questions = []
    q_num = 1

    basic_wrong = []
    advanced_wrong = []

    if wrong_basic_str and str(wrong_basic_str).strip() and str(wrong_basic_str).strip() != "nan":
        basic_wrong = [int(x.strip()) for x in str(wrong_basic_str).split(",") if x.strip().isdigit()]
    if wrong_advanced_str and str(wrong_advanced_str).strip() and str(wrong_advanced_str).strip() != "nan":
        advanced_wrong = [int(x.strip()) for x in str(wrong_advanced_str).split(",") if x.strip().isdigit()]

    for num in basic_wrong:
        if num in basic_question_map:
            questions.append(f"{q_num}. {basic_question_map[num]}")
            q_num += 1

    for num in advanced_wrong:
        if num in advanced_question_map:
            questions.append(f"{q_num}. {advanced_question_map[num]}")
            q_num += 1

    fallback_questions = [
        f"{q_num}. 請寫出『光合作用需要葉綠體』的實驗證據。（提示：白斑葉實驗）",
        f"{q_num + 1}. 請說明：人體為什麼需要攝取纖維素，但它卻無法提供熱量？",
        f"{q_num + 2}. 請比較『原核生物』與『真核生物』在細胞構造上的主要差異。"
    ]
    for fallback in fallback_questions:
        if len(questions) < min_q:
            questions.append(fallback)
            q_num += 1

    return questions


# =============================================
# 產出 PDF 函數
# =============================================

def create_student_pdf(row, output_dir, class_name_val="9年01班"):
    font_name = "CustomFont"

    title_style = ParagraphStyle("DocTitle", fontName=font_name, fontSize=17, leading=21,
                                  textColor=colors.HexColor("#1A5276"), alignment=1, spaceAfter=12)
    section_style = ParagraphStyle("SectionHeading", fontName=font_name, fontSize=12.5, leading=17,
                                    textColor=colors.HexColor("#1A5276"), spaceBefore=12, spaceAfter=6)
    body_style = ParagraphStyle("BodyText", fontName=font_name, fontSize=10.5, leading=15.5,
                                 textColor=colors.HexColor("#2C3E50"), spaceBefore=3, spaceAfter=4)
    advice_style = ParagraphStyle("AdviceText", fontName=font_name, fontSize=10.5, leading=15.5,
                                   textColor=colors.HexColor("#0E6251"), spaceBefore=3, spaceAfter=3)
    q_style = ParagraphStyle("QuestionText", fontName=font_name, fontSize=10.5, leading=15.5,
                              textColor=colors.HexColor("#1A1A1A"), spaceBefore=6, spaceAfter=3)

    pdf_filename = os.path.join(output_dir, f"{row['座號']}_{row['姓名']}_生物科診斷報告.pdf")
    doc = SimpleDocTemplate(pdf_filename, pagesize=A4, rightMargin=35, leftMargin=35, topMargin=35, bottomMargin=35)
    story = []

    story.append(Paragraph("<b>《生物科第2回複習考》個人化診斷報告</b>", title_style))
    story.append(Spacer(1, 3))

    wrong_display = f"基礎題：{row['基礎錯題']} | 精熟題：{row['精熟錯題']}"
    if pd.isna(row.get("基礎錯題", "")) and pd.isna(row.get("精熟錯題", "")):
        wrong_display = "無（全對！）"

    header_data = [
        [Paragraph(f"<b>班級：</b>{class_name_val}", body_style), Paragraph(f"<b>座號：</b>{row['座號']} 號", body_style)],
        [Paragraph(f"<b>姓名：</b>{row['姓名']}", body_style),
         Paragraph(f"<b>測驗得分：<font color='#E74C3C'><b>{row['總分']} 分</b></font></b>", body_style)],
        [Paragraph(f"<b>錯題分布：</b>{wrong_display}", body_style), ""]
    ]
    header_table = Table(header_data, colWidths=[245, 245])
    header_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F2F9FF")),
        ("BOX", (0, 0), (-1, -1), 1.2, colors.HexColor("#3498DB")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("SPAN", (0, 2), (1, 2))
    ]))
    story.append(header_table)
    story.append(Spacer(1, 8))

    story.append(Paragraph("<b>一、 誘答機制與核心考點拆解</b>", section_style))
    analysis_text = row.get("分析文字", "（請在 Excel 中填入分析文字）")
    story.append(Paragraph(str(analysis_text) if pd.notna(analysis_text) else "（請在 Excel 中填入分析文字）", body_style))
    story.append(Spacer(1, 6))

    story.append(Paragraph("<b>二、 思維路徑與學習認知診斷</b>", section_style))
    mindset_text = row.get("思維診斷", "（請在 Excel 中填入思維診斷）")
    story.append(Paragraph(str(mindset_text) if pd.notna(mindset_text) else "（請在 Excel 中填入思維診斷）", body_style))
    story.append(Spacer(1, 6))

    story.append(Paragraph("<b>三、 老師給你的暖心引導與挑戰</b>", section_style))
    advice_text = row.get("暖心引導", "（請在 Excel 中填入暖心引導）")
    advice_content = [[Paragraph(f"「{advice_text}」" if pd.notna(advice_text) else "（請在 Excel 中填入暖心引導）", advice_style)]]
    advice_table = Table(advice_content, colWidths=[490])
    advice_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#E9F7EF")),
        ("BOX", (0, 0), (-1, -1), 1.2, colors.HexColor("#27AE60")),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12)
    ]))
    story.append(KeepTogether([advice_table]))
    story.append(Spacer(1, 8))

    story.append(PageBreak())
    story.append(Paragraph(f"<b>📝 {row['姓名']} 同學 專屬補考練習卷</b>", title_style))
    story.append(Spacer(1, 6))
    story.append(Paragraph(f"班級：{class_name_val}  姓名：{row['姓名']}  座號：{row['座號']}", body_style))
    story.append(Spacer(1, 8))

    questions = generate_practice_questions(row.get("基礎錯題", ""), row.get("精熟錯題", ""), min_questions)

    basic_wrong_display = row.get('基礎錯題', '無')
    advanced_wrong_display = row.get('精熟錯題', '無')
    if pd.isna(basic_wrong_display):
        basic_wrong_display = '無'
    if pd.isna(advanced_wrong_display):
        advanced_wrong_display = '無'
    source_text = f"📌 本卷對應你的錯題：基礎題 {basic_wrong_display} | 精熟題 {advanced_wrong_display}"
    story.append(Paragraph(source_text, body_style))
    story.append(Spacer(1, 8))

    for q in questions:
        story.append(Paragraph(q, q_style))
        story.append(Spacer(1, 4))

    story.append(Spacer(1, 12))
    story.append(Paragraph("<b>📌 答案欄（請寫下你的答案）：</b>", body_style))
    for i in range(1, len(questions) + 1):
        story.append(Paragraph(f"第 {i} 題：______", body_style))
        story.append(Spacer(1, 2))

    doc.build(story)
    return pdf_filename


# =============================================
# 執行產生
# =============================================

if "df" in st.session_state:
    if st.button("🚀 產生所有學生的診斷報告", type="primary", use_container_width=True):
        df = st.session_state.df
        with st.spinner("⏳ 正在生成 PDF，請稍候（約 10~30 秒）..."):
            output_dir = tempfile.mkdtemp()
            class_name_val = class_name if show_class else ""
            success_count = 0

            progress_bar = st.progress(0)
            status_text = st.empty()
            total = len(df)

            for position, (idx, row) in enumerate(df.iterrows()):
                status_text.text(f"正在處理：{row['姓名']} ({position + 1}/{total})")
                try:
                    create_student_pdf(row, output_dir, class_name_val)
                    success_count += 1
                except Exception as e:
                    st.warning(f"⚠️ {row['姓名']} 產生失敗：{str(e)}")
                progress_bar.progress((position + 1) / total)

            status_text.text("✅ 全部完成！")

            zip_path = os.path.join(tempfile.gettempdir(), "診斷報告_全學生.zip")
            with zipfile.ZipFile(zip_path, "w") as zipf:
                for file in os.listdir(output_dir):
                    zipf.write(os.path.join(output_dir, file), file)

            st.session_state.zip_path = zip_path
            st.session_state.output_dir = output_dir
            st.session_state.success_count = success_count

            st.success(f"🎉 成功產生 {success_count}/{total} 位學生的個人化診斷報告！")
else:
    st.info("ℹ️ 請先在步驟一上傳 Excel 檔案，才能產生診斷報告")

# =============================================
# 下載按鈕
# =============================================

if "zip_path" in st.session_state:
    st.markdown("## 📥 步驟三：下載報告")

    with open(st.session_state.zip_path, "rb") as f:
        st.download_button(
            label="📥 下載所有診斷報告 (ZIP 壓縮檔)",
            data=f,
            file_name="診斷報告_全學生.zip",
            mime="application/zip",
            use_container_width=True
        )

    st.info(f"📂 ZIP 檔內含 {st.session_state.success_count} 份 PDF 報告，每份包含診斷分析 + 個人化補考卷")

# =============================================
# 寄送 Email（選用）
# =============================================

st.markdown("## 📧 步驟四：寄送 Email 給學生（選用）")

sender_email = ""
sender_password = ""
subject = "🧬 你的生物科個人化診斷報告"
email_body_template = (
    "親愛的 {姓名} 同學：\n\n"
    "這是你的生物科個人化診斷報告與專屬補考練習卷，請仔細閱讀。\n\n"
    "裡面有針對你的錯題設計的專屬引導與練習題，記得在下次上課前完成喔！\n\n"
    "加油！🔥\n\n老師 敬上"
)

with st.expander("⚙️ 設定 Gmail 寄信（點擊展開）"):
    st.warning("""
    **使用前注意事項**：
    1. 必須啟用 Gmail 的「兩步驟驗證」
    2. 必須產生「應用程式密碼」（16碼）
    3. [點此前往 Google 應用程式密碼設定](https://myaccount.google.com/apppasswords)
    """)

    col1, col2 = st.columns(2)
    with col1:
        sender_email = st.text_input("寄件者 Email（Gmail）", placeholder="your_email@gmail.com")
    with col2:
        sender_password = st.text_input("應用程式密碼", type="password", placeholder="請輸入 16 位密碼")

    subject = st.text_input("信件主旨", value=subject)

    email_body_template = st.text_area(
        "信件內文（可使用 {姓名} 作為個人化變數）",
        value=email_body_template,
        height=150
    )

    if st.button("📨 測試寄信（寄給自己）", type="secondary"):
        if not sender_email or not sender_password:
            st.error("❌ 請先填寫寄件者 Email 和應用程式密碼")
        else:
            try:
                server = smtplib.SMTP("smtp.gmail.com", 587)
                server.starttls()
                server.login(sender_email, sender_password)
                server.quit()
                st.success("✅ 連線測試成功！可以寄信了！")
            except Exception as e:
                st.error(f"❌ 連線失敗：{e}")
                st.info("請確認應用程式密碼是否正確，或檢查 Gmail 兩步驟驗證是否啟用")

if st.button("📧 寄送 Email 給所有學生（有 Email 者）", type="secondary", use_container_width=True):
    if not sender_email or not sender_password:
        st.error("❌ 請先填寫寄件者 Email 和應用程式密碼")
    elif "df" not in st.session_state:
        st.error("❌ 尚未上傳學生資料")
    elif "Email" not in st.session_state.df.columns:
        st.error("❌ Excel 中沒有「Email」欄位，無法寄信")
    elif "output_dir" not in st.session_state:
        st.error("❌ 請先產生診斷報告，才能寄送附件")
    else:
        df = st.session_state.df
        with st.spinner("⏳ 正在寄送 Email..."):
            try:
                server = smtplib.SMTP("smtp.gmail.com", 587)
                server.starttls()
                server.login(sender_email, sender_password)

                success_count = 0
                fail_count = 0
                progress_bar = st.progress(0)
                status_text = st.empty()
                total = len(df)

                for position, (idx, row) in enumerate(df.iterrows()):
                    student_email = row.get("Email", "")
                    if pd.isna(student_email) or not str(student_email).strip():
                        fail_count += 1
                        continue

                    status_text.text(f"正在寄送：{row['姓名']} -> {student_email}")

                    try:
                        msg = MIMEMultipart()
                        msg["From"] = sender_email
                        msg["To"] = str(student_email).strip()
                        msg["Subject"] = subject

                        body = email_body_template.format(姓名=row["姓名"])
                        msg.attach(MIMEText(body, "plain", "utf-8"))

                        pdf_file = os.path.join(st.session_state.output_dir, f"{row['座號']}_{row['姓名']}_生物科診斷報告.pdf")
                        if os.path.exists(pdf_file):
                            with open(pdf_file, "rb") as f:
                                part = MIMEBase("application", "octet-stream")
                                part.set_payload(f.read())
                                encoders.encode_base64(part)
                                part.add_header("Content-Disposition", f"attachment; filename={row['姓名']}_診斷報告.pdf")
                                msg.attach(part)
                            server.send_message(msg)
                            success_count += 1
                        else:
                            st.warning(f"⚠️ {row['姓名']} 的 PDF 檔案不存在，跳過寄送")
                            fail_count += 1
                    except Exception as e:
                        st.warning(f"⚠️ {row['姓名']} 寄送失敗：{str(e)}")
                        fail_count += 1

                    progress_bar.progress((position + 1) / total)

                server.quit()
                status_text.text("✅ 寄送完成！")
                st.success(f"✅ 成功寄送 {success_count} 封信件，{fail_count} 位學生無 Email 或寄送失敗")

            except Exception as e:
                st.error(f"❌ 寄送過程發生錯誤：{e}")
                st.info("請檢查 Email 設定或網路連線")

# =============================================
# 頁尾
# =============================================

st.markdown("---")
st.caption("🧬 生物科個人化診斷報告產生器 v2.0 | 資料僅在本次作業階段處理，不會儲存")
