# =====================================================================
# 🧬 生物科個人化診斷報告產生器 v2.0（Streamlit 版）
# =====================================================================

import streamlit as st
import pandas as pd
import os
import shutil
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

# =============================================
# 頁面設定
# =============================================

st.set_page_config(page_title="生物科診斷報告產生器", layout="wide")

st.title("🧬 生物科個人化診斷報告產生器")
st.markdown("上傳學生資料 Excel，一鍵產生所有學生的個人化診斷報告 + 補考卷")

# =============================================
# 上傳字型
# =============================================

st.sidebar.header("📁 上傳字型檔")
uploaded_font = st.sidebar.file_uploader("上傳 .ttf 或 .ttc 字型檔", type=["ttf", "ttc"])

if uploaded_font:
    font_path = os.path.join(tempfile.gettempdir(), uploaded_font.name)
    with open(font_path, "wb") as f:
        f.write(uploaded_font.getbuffer())
    pdfmetrics.registerFont(TTFont("CustomFont", font_path))
    st.sidebar.success("✅ 字型註冊成功！")
else:
    st.sidebar.warning("⚠️ 請上傳字型檔（推薦：NotoSansTC-Regular.ttf）")

# =============================================
# 上傳 Excel 資料
# =============================================

st.header("📤 步驟一：上傳學生資料")

uploaded_file = st.file_uploader("上傳 Excel 檔（.xlsx）", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.success(f"✅ 成功讀取 {len(df)} 位學生資料！")

    # 顯示資料預覽
    with st.expander("📋 資料預覽（前5筆）"):
        st.dataframe(df.head(5))

    # =============================================
    # 錯誤題庫（對應補考卷）
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

    def generate_practice_questions(wrong_basic_str, wrong_advanced_str):
        questions = []
        q_num = 1

        basic_wrong = [int(x.strip()) for x in str(wrong_basic_str).split(',') if x.strip().isdigit()]
        advanced_wrong = [int(x.strip()) for x in str(wrong_advanced_str).split(',') if x.strip().isdigit()]

        for num in basic_wrong:
            if num in basic_question_map:
                questions.append(f"{q_num}. {basic_question_map[num]}")
                q_num += 1

        for num in advanced_wrong:
            if num in advanced_question_map:
                questions.append(f"{q_num}. {advanced_question_map[num]}")
                q_num += 1

        if len(questions) < 5:
            questions.append(f"{q_num}. 請寫出『光合作用需要葉綠體』的實驗證據。")
            q_num += 1
        if len(questions) < 5:
            questions.append(f"{q_num}. 請說明人體為什麼需要攝取纖維素，但它卻無法提供熱量？")
            q_num += 1

        return questions

    # =============================================
    # 產生 PDF
    # =============================================

    if st.button("🚀 產生所有學生的診斷報告", type="primary"):
        if not uploaded_font:
            st.error("❌ 請先上傳字型檔！")
        else:
            with st.spinner("⏳ 正在生成 PDF，請稍候..."):
                output_dir = tempfile.mkdtemp()

                # 樣式設定
                font_name = "CustomFont"
                title_style = ParagraphStyle('DocTitle', fontName=font_name, fontSize=17, leading=21, textColor=colors.HexColor("#1A5276"), alignment=1, spaceAfter=12)
                section_style = ParagraphStyle('SectionHeading', fontName=font_name, fontSize=12.5, leading=17, textColor=colors.HexColor("#1A5276"), spaceBefore=12, spaceAfter=6)
                body_style = ParagraphStyle('BodyText', fontName=font_name, fontSize=10.5, leading=15.5, textColor=colors.HexColor("#2C3E50"), spaceBefore=3, spaceAfter=4)
                advice_style = ParagraphStyle('AdviceText', fontName=font_name, fontSize=10.5, leading=15.5, textColor=colors.HexColor("#0E6251"), spaceBefore=3, spaceAfter=3)
                q_style = ParagraphStyle('QuestionText', fontName=font_name, fontSize=10.5, leading=15.5, textColor=colors.HexColor("#1A1A1A"), spaceBefore=6, spaceAfter=3)

                progress_bar = st.progress(0)
                status_text = st.empty()

                for idx, row in df.iterrows():
                    status_text.text(f"正在處理：{row['姓名']} ({idx+1}/{len(df)})")

                    pdf_filename = os.path.join(output_dir, f"{row['座號']}_{row['姓名']}_生物科診斷報告.pdf")
                    doc = SimpleDocTemplate(pdf_filename, pagesize=A4, rightMargin=35, leftMargin=35, topMargin=35, bottomMargin=35)
                    story = []

                    # 標題
                    story.append(Paragraph("<b>《生物科第2回複習考》個人化診斷報告</b>", title_style))
                    story.append(Spacer(1, 3))

                    # 基本資料
                    wrong_display = f"基礎題：{row['基礎錯題']} | 精熟題：{row['精熟錯題']}"
                    if pd.isna(row['基礎錯題']) and pd.isna(row['精熟錯題']):
                        wrong_display = "無（全對！）"

                    header_data = [
                        [Paragraph(f"<b>班級：</b>9年01班", body_style), Paragraph(f"<b>座號：</b>{row['座號']} 號", body_style)],
                        [Paragraph(f"<b>姓名：</b>{row['姓名']}", body_style), Paragraph(f"<b>測驗得分：<font color='#E74C3C'><b>{row['總分']} 分</b></font></b>", body_style)],
                        [Paragraph(f"<b>錯題分布：</b>{wrong_display}", body_style), ""]
                    ]
                    header_table = Table(header_data, colWidths=[245, 245])
                    header_table.setStyle(TableStyle([
                        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F2F9FF")),
                        ('BOX', (0,0), (-1,-1), 1.2, colors.HexColor("#3498DB")),
                        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                        ('TOPPADDING', (0,0), (-1,-1), 6),
                        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
                        ('LEFTPADDING', (0,0), (-1,-1), 12),
                        ('SPAN', (0,2), (1,2))
                    ]))
                    story.append(header_table)
                    story.append(Spacer(1, 8))

                    # 考點拆解
                    story.append(Paragraph("<b>一、 誘答機制與核心考點拆解</b>", section_style))
                    analysis_text = row.get('分析文字', '（請在 Excel 中填入分析文字）')
                    story.append(Paragraph(str(analysis_text), body_style))
                    story.append(Spacer(1, 6))

                    # 思維診斷
                    story.append(Paragraph("<b>二、 思維路徑與學習認知診斷</b>", section_style))
                    mindset_text = row.get('思維診斷', '（請在 Excel 中填入思維診斷）')
                    story.append(Paragraph(str(mindset_text), body_style))
                    story.append(Spacer(1, 6))

                    # 暖心引導
                    story.append(Paragraph("<b>三、 老師給你的暖心引導與挑戰</b>", section_style))
                    advice_text = row.get('暖心引導', '（請在 Excel 中填入暖心引導）')
                    advice_content = [[Paragraph(f"「{advice_text}」", advice_style)]]
                    advice_table = Table(advice_content, colWidths=[490])
                    advice_table.setStyle(TableStyle([
                        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#E9F7EF")),
                        ('BOX', (0,0), (-1,-1), 1.2, colors.HexColor("#27AE60")),
                        ('TOPPADDING', (0,0), (-1,-1), 10),
                        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
                        ('LEFTPADDING', (0,0), (-1,-1), 12),
                        ('RIGHTPADDING', (0,0), (-1,-1), 12)
                    ]))
                    story.append(KeepTogether([advice_table]))
                    story.append(Spacer(1, 8))

                    # 補考卷
                    story.append(PageBreak())
                    story.append(Paragraph(f"<b>📝 {row['姓名']} 同學 專屬補考練習卷</b>", title_style))
                    story.append(Spacer(1, 6))
                    story.append(Paragraph(f"班級：901  姓名：{row['姓名']}  座號：{row['座號']}", body_style))
                    story.append(Spacer(1, 8))

                    questions = generate_practice_questions(
                        row.get('基礎錯題', ''),
                        row.get('精熟錯題', '')
                    )

                    source_text = f"📌 本卷對應你的錯題：基礎題 {row.get('基礎錯題', '無')} | 精熟題 {row.get('精熟錯題', '無')}"
                    story.append(Paragraph(source_text, body_style))
                    story.append(Spacer(1, 8))

                    for q in questions:
                        story.append(Paragraph(q, q_style))
                        story.append(Spacer(1, 4))

                    story.append(Spacer(1, 12))
                    story.append(Paragraph("<b>📌 答案欄（請寫下你的答案）：</b>", body_style))
                    for i in range(1, len(questions)+1):
                        story.append(Paragraph(f"第 {i} 題：______", body_style))
                        story.append(Spacer(1, 2))

                    doc.build(story)

                    progress_bar.progress((idx + 1) / len(df))

                # 打包 ZIP
                status_text.text("📦 正在打包 ZIP 檔...")
                zip_path = os.path.join(tempfile.gettempdir(), "診斷報告_全學生.zip")
                with zipfile.ZipFile(zip_path, 'w') as zipf:
                    for file in os.listdir(output_dir):
                        zipf.write(os.path.join(output_dir, file), file)

                status_text.text("✅ 完成！")

                # 提供下載按鈕
                with open(zip_path, 'rb') as f:
                    st.download_button(
                        label="📥 下載所有診斷報告 (ZIP)",
                        data=f,
                        file_name="診斷報告_全學生.zip",
                        mime="application/zip"
                    )

                st.success(f"🎉 成功產生 {len(df)} 位學生的個人化診斷報告！")

                # =============================================
                # 寄送 Email
                # =============================================

                st.subheader("📧 步驟二：寄送 Email 給學生")

                with st.expander("⚙️ 設定寄信選項"):
                    sender_email = st.text_input("寄件者 Email（Gmail）", placeholder="your_email@gmail.com")
                    sender_password = st.text_input("應用程式密碼", type="password", placeholder="16碼應用程式密碼")
                    subject = st.text_input("信件主旨", value="🧬 你的生物科個人化診斷報告")

                    email_body_template = st.text_area(
                        "信件內文（可使用 {姓名} 作為變數）",
                        value="""親愛的 {姓名} 同學：

這是你的生物科個人化診斷報告與專屬補考練習卷，請仔細閱讀。

裡面有針對你的錯題設計的專屬引導與練習題，記得在下次上課前完成喔！

加油！🔥

老師 敬上"""
                    )

                if st.button("📧 寄送 Email 給所有學生", type="secondary"):
                    if not sender_email or not sender_password:
                        st.error("❌ 請填寫寄件者 Email 和應用程式密碼")
                    else:
                        with st.spinner("⏳ 正在寄送 Email..."):
                            try:
                                server = smtplib.SMTP("smtp.gmail.com", 587)
                                server.starttls()
                                server.login(sender_email, sender_password)

                                success_count = 0
                                for idx, row in df.iterrows():
                                    student_email = row.get('Email', '')
                                    if pd.isna(student_email) or not student_email:
                                        continue

                                    msg = MIMEMultipart()
                                    msg['From'] = sender_email
                                    msg['To'] = student_email
                                    msg['Subject'] = subject

                                    body = email_body_template.format(姓名=row['姓名'])
                                    msg.attach(MIMEText(body, 'plain'))

                                    pdf_file = os.path.join(output_dir, f"{row['座號']}_{row['姓名']}_生物科診斷報告.pdf")
                                    if os.path.exists(pdf_file):
                                        with open(pdf_file, 'rb') as f:
                                            part = MIMEBase('application', 'octet-stream')
                                            part.set_payload(f.read())
                                            encoders.encode_base64(part)
                                            part.add_header('Content-Disposition', f'attachment; filename={row["姓名"]}_診斷報告.pdf')
                                            msg.attach(part)

                                        server.send_message(msg)
                                        success_count += 1

                                server.quit()
                                st.success(f"✅ 成功寄送 {success_count} 封信件！")
                            except Exception as e:
                                st.error(f"❌ 寄送失敗：{e}")
