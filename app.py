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

# =============================================
# 頁面設定
# =============================================

st.set_page_config(
    page_title="生物科診斷報告產生器",
    page_icon="🧬",
    layout="wide"
)

# =============================================
# 側邊欄：使用教學（永遠顯示）
# =============================================

with st.sidebar:
    st.markdown("## 📖 使用教學")
    st.markdown("---")
    
    with st.expander("📌 第一步：準備資料", expanded=True):
        st.markdown(
            """
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
            """
        )
    
    with st.expander("✍️ 第二步：填入診斷文字"):
        st.markdown(
            """
            **分析文字**：可以用 HTML 標籤，例如：
