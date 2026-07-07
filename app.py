import os
import time
import smtplib
import shutil
import tempfile
import urllib.request
import pandas as pd
import streamlit as st
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# 設定頁面網頁標題
st.set_page_config(page_title="生物科 AI 診斷與自動發信系統", page_icon="🔬", layout="wide")

# 1. 雲端下載台灣開源「芫荽手寫體」字型 (確保任何環境都不亂碼)
FONT_PATH = "Iansui-Regular.ttf"
FONT_NAME = "IansuiHandwriting"

@st.cache_resource
def load_font():
    if not os.path.exists(FONT_PATH):
        try:
            with st.spinner("🔄 正在從雲端下載『台灣教育硬筆手寫體 (芫荽體)』..."):
                url = "https://github.com/ButKo/iansui/raw/main/releases/Iansui-Regular.ttf"
                urllib.request.urlretrieve(url, FONT_PATH)
        except Exception as e:
            st.error(f"⚠️ 字型下載失敗，將使用系統預設字型。錯誤資訊: {e}")
            return "Helvetica"
    
    try:
        pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))
        return FONT_NAME
    except Exception as e:
        st.error(f"⚠️ 字型註冊失敗: {e}")
        return "Helvetica"

active_font = load_font()

# 2. 11 位學生超深度教學診斷核心資料庫
STUDENTS_DB = {
    "10": {
        "name": "陳昱O", "score": "89.5", "wrong": "基礎題：4, 13, 17 (精熟進階題全對！)",
        "analysis": "<b>1. 基礎題 4【養分與熱量迷思】：</b>出題老師在這裡埋下了「攝取食物就是為了獲取熱量」的直覺陷阱！纖維素雖然屬於多醣類，但人體缺乏分解它的酵素，因此絕對「無法提供熱量（能量）」；考生容易把它和澱粉混淆，忘記人體攝取纖維素的真正目的是為了「刺激腸道蠕動、幫助排遺」。<br/><br/>"
                    "<b>2. 基礎題 13【生物攝食構造與適應】：</b>這題考驗的是「單細胞 vs. 多細胞」生物的攝食機制差異。容易落入陷阱是因為習慣以多細胞動物的思維（牙齒、管狀口器、喙）去推想所有生物；變形蟲屬於單細胞生物，必須利用細胞膜延伸進行「吞噬作用」將食物包覆並形成食泡。<br/><br/>"
                    "<b>3. 基礎題 17【消化道與消化腺機能陷阱】：</b>這是國中生物最經典的「膽囊騙局」！出題老師利用「膽汁在膽囊」的字面聯想，誘導學生誤以為膽囊會「分泌」膽汁。事實上，膽汁是「肝臟」製造分泌的，膽囊只是一個負責儲存的「暫存倉庫」，本身絕對不具有分泌消化液的功能！",
        "mindset": "【邏輯大師，細節迷路】昱O在需要多步驟推理、實驗變因分析與圖表判讀的「精熟題（進階題）」中，竟然能做到完全無失分（全對）！這代表你的理科邏輯思維與科學素養已經屬於極其頂尖的水準。目前思維上唯一的結構性盲點，在於對「基礎生理名詞與器官機能」的記憶不夠精確，容易被生活中的直覺（例如：吃纖維＝吸收熱量、膽囊＝分泌膽汁）給取代。只要把這幾個考卷中的「經典防呆關鍵字」重新鎖死，你就是無懈可擊的滿分學霸！",
        "advice": "昱O啊，先用力給自己的大腦鼓鼓掌！考卷最難的精熟進階題你全部秒殺全對，你的理科思考邏輯真的超強，老師太佩服你了！不過下次寫考卷記得把班級座號寫對，不然老師差點找不到這位隱形高手啦！針對這次錯的三個小細節，老師不直接餵你答案，請你動動腦想一想這三個問題：(1) 牛和羊吃草可以長得胖嘟嘟，但我們人類瘋狂吃蔬菜（纖維素）卻能幫助排便與減肥，你覺得我們肚子裡是不是少了什麼能『把纖維素變成熱量』的秘密武器？那我們吃它到底是在幫腸道做什麼運動？(2) 如果今天你是一隻沒有手腳、也沒有牙齒的單細胞『變形蟲』，看到前面有一塊美味的食物，你要怎麼用自己的『身體（細胞膜）』把它給「抱」進來變成食泡？(3) 肝臟跟膽囊就像是一對合作無間的開店夥伴，一個負責『工廠製造』，一個負責『倉庫囤貨』。下次看到膽囊時，請幫老師想一想：倉庫老闆自己會不會生產產品？膽汁到底是誰製造出來的？"
    },
    "22": {
        "name": "梁容O", "score": "97.5", "wrong": "精熟題：11 (全卷 32 題僅錯此單一題型！)",
        "analysis": "<b>1. 精熟題 11【原核與真核生物胞器對比】：</b>此題是出題老師精心設計的「高階概念誘答題」。題目圖示中行光合作用的橢圓形包膜胞器為「葉綠體」。考生易因「藍綠菌也是水中的光合生物」而誤選；然而核心破體點在於分類學的階層界線——藍綠菌屬於「原核生物」，細胞內沒有核膜，更絕對不具有葉綠體等膜狀包圍的胞器（其光合作用是仰賴分散在細胞質中的葉綠素）。因此，這個有膜狀構造包圍的葉綠體，只能在屬於真核生物的「矽藻」細胞中被發現。",
        "mindset": "【頂尖學霸，僅缺最後一道防呆鎖】容O在全卷 32 題中，基礎題 20 題全對、精熟題 11 題全對，拿下全班極為頂尖的 97.5 分！解題邏輯、專注力與生理概念已近乎極致。目前唯一的思維斷裂點在於原核生物的胞器特化定義：在推論時，光合作用的生理特徵覆蓋了細胞學的構造分類。只要把「原核生物＝無核膜、無膜狀胞器」列為最高優先權的反饋條件，就能達到完美的 100 分。",
        "advice": "容O，老師真的要給你最熱烈的掌聲！全卷 30 幾道題目你竟然只錯一題，考出 97.5 分的超級頂尖高分！你對生物消化系統、酵素實驗的理解已經達到學霸級別。這唯一失手的一題，是高中生也常常被騙的王牌陷阱喔！記住老師送你的防呆金句：『有葉綠素不等於有葉綠體；藍綠菌是原核界，絕對沒有葉綠體！』只要把這條界線鎖死，下次會考的滿分榜首絕對非你莫屬！繼續保持這個驚人的氣勢！"
    },
    "30": {
        "name": "鄭采O", "score": "91.5", "wrong": "基礎題：13；精熟題：5, 11",
        "analysis": "<b>1. 基礎題 13【生物攝食構造與適應】：</b>考生容易以巨觀多細胞動物的直覺去推測單細胞生物，忽略變形蟲必須藉由「細胞膜延伸」進行吞噬作用形成食泡來攝食環境中的食物；豹屬於哺乳類，是用「牙齒」咀嚼而非口器。<br/><br/>"
                    "<b>2. 精熟題 5【大分子養分消化順序】：</b>出題老師以圖表隱藏三大養分的消化起點。澱粉（馬鈴薯）最早在口腔被唾液分解；蛋白質（豆腐）其次在胃被胃液分解；脂質（橄欖油）最晚在小腸藉由膽汁乳化與胰液分解。如果沒有把食物種類與消化管順序串接，極易落入配對陷阱。<br/><br/>"
                    "<b>3. 精熟題 11【原核與真核胞器差異】：</b>此題為經典的「葉綠素不等於葉綠體」迷思陷阱。藍綠菌雖能行光合作用，但屬於原核生物，細胞內沒有核膜，也絕對沒有葉綠體等膜狀胞器；橢圓形的葉綠體只能在真核藻類（如矽藻）中發現。",
        "mindset": "【高手卡關在細節對比】采O的基礎極為扎實，能拿下 91.5 分的高分，說明對核心概念掌握度極佳。目前思維上唯一的結構性盲點在於「跨領域概念的交叉對比」（如食物分類對應消化管時序、原核真核的構造有無）。只要把「消化時間軸」與「胞器防呆口訣」建立起來，就能輕取滿分。",
        "advice": "采O！你的生物底子真的很強，91.5 分絕對是班上的領頭羊之一！不過老師發現你這次稍微被出題老師的『相似詞陷阱』給忽悠了喔。請你挑戰這兩個進階問題：(1) 如果我們把人體消化道當成一條加工輸送帶（口腔 -> 胃 -> 小腸），請你幫老師把今天吃的『馬鈴薯、豆腐、橄欖油』依序放上輸送帶，誰會最先遭遇分解酵素？誰又要等到最後一關才被處理？(2) 藍綠菌就像是用原始的帳篷露營者（原核生物），而矽藻是住進豪華別墅的現代人（真核生物）。想一想，藍綠菌雖然能行光合作用，但它肚子裡到底有沒有被膜包起來的『葉綠體房間』呢？"
    },
    "28": {
        "name": "廖禹O", "score": "77.5", "wrong": "基礎題：4, 13, 14, 16, 17；精熟題：1, 8",
        "analysis": "<b>1. 基礎題 4、16【不供能養分迷思】：</b>容易混淆「熱量」與「生理機能」的關聯。纖維素（多醣類）與維生素都完全無法在人體內產生熱量（能量）；吃蔬菜（纖維素）的唯一力學目標是刺激腸道蠕動、幫助排遺。<br/><br/>"
                    "<b>2. 基礎題 14【生物能量獲取方式】：</b>出題老師用「捕蟲植物（豬籠草、捕蠅草）」混淆考生的營養學分類。捕蟲植物本質上是「自營生產者」，能藉由光合作用自行製造葡萄糖；吃昆蟲只是為了補充土壤缺乏的「氮元素」，這與異營動物完全依賴攝食獲得能量有本質上的大逆轉。<br/><br/>"
                    "<b>3. 基礎題 17、精熟題 1【消化道機能陷阱】：</b>深陷「膽囊分泌膽汁」與「大腸分泌消化液」的直覺錯誤。膽囊只負責「儲存」，肝臟才是「製造與分泌」膽汁的器官；大腸主要負責「吸收水分與形成糞便」，絕不分泌任何消化液。此外，胃內部因含有胃酸（鹽酸），環境為「強酸性」。<br/><br/>"
                    "<b>4. 精熟題 8【酵素圖表判讀】：</b>將圖表橫軸（溫度跨度）與縱軸（活性高低）混淆。曲線在橫軸延伸的寬度越寬，代表該酵素能忍受的溫度範圍較廣（溫度的容忍程度高）。",
        "mindset": "【概念混淆與器官機能斷裂】禹O目前在「六大養分生理功能」與「人體消化器官分工」的記憶鏈條出現嚴重斷裂。對科學圖表的「最佳值（頂點）」與「容忍度（底軸跨度）」缺乏直覺對應，需要透過親自繪圖與顏色標註來重建腦內知識地圖。",
        "advice": "禹O，這次考卷是不是覺得很多題目長得都很像，讓你眼花撩亂了？別擔心，這代表你的觀念正處於『重組期』！老師不給你標準答案，請你跟著老師的提示自己破解：(1) 豬籠草雖然看起來很兇會吃蟲，但它在大自然裡的真實身分到底是『自己行光合作用製造養分』的生產者，還是『只能到處吃東西』的消費者？它吃蟲到底是在補充什麼神祕元素？(2) 請你在紙上畫出肝臟、膽囊和大腸。記住老師的防呆口訣：『膽囊只是暫存倉庫，不生產產品；大腸是海綿只負責吸水，不吐消化液』。(3) 如果有兩支酵素，A 酵素只能在 35~40度C 活著，B 酵素能在 10~80度C 都活得很好，你覺得在科學圖表的曲線跨度上，誰是那個『對環境容忍度超高』的霸主？"
    },
    "11": {
        "name": "游諺O", "score": "88.0", "wrong": "基礎題：4, 13；精熟題：8, 11",
        "analysis": "<b>1. 基礎題 4【纖維素迷思】：</b>人體消化道缺乏分解纖維素的酵素，因此吃蔬菜不是為了化學產熱，而是利用其高纖維物理特性來刺激腸道蠕動。<br/><br/>"
                    "<b>2. 基礎題 13【攝食適應】：</b>忽略變形蟲為單細胞生物，必須以細胞膜延伸進行「吞噬作用」形成食泡來攝食環境食物。<br/><br/>"
                    "<b>3. 精熟題 8【溫度容忍度圖表】：</b>未能精準判讀科學圖表的橫軸物理意義。當曲線在橫軸（溫度）覆蓋的範圍越寬，代表該酵素在波動的溫度下較不易失活，即對溫度的容忍程度較高。<br/><br/>"
                    "<b>4. 精熟題 11【原核胞器辨識】：</b>被「能行光合作用的生物必有葉綠體」的先入為主概念誘導。藍綠菌為原核生物，僅具有葉綠素（光合色素），無葉綠體；完整的葉綠體只存在於真核藻類（如矽藻）與植物中。",
        "mindset": "【觀念清晰，定義精確度稍欠】諺O的整體學習狀況良好（88分），主要錯在典型的考點陷阱。概念連結的精密解析度不足，在腦中將「光合作用」與「葉綠體」劃上了絕對等號，忽略了「生物分類（原核/真核）」這個最高層級的過濾條件。",
        "advice": "諺O！能考到 88 分代表你的生物基礎相當優秀，老師為你拍拍手！你現在離 90 分以上的高手區只差最後一哩路，這幾題錯的都是全班最經典的『易錯陷阱』。來，幫老師思考兩個小難關：(1) 我們常說『有葉綠體的生物才能行光合作用』，這句話在遇到『藍綠菌』這位原始的原核生物時，為什麼會破功？IT體內操縱的到底是什麼『色素』在抓太陽光？(2) 看實驗圖表時，請把手指放在橫軸（底線）上。如果一條曲線的底邊長得超寬，從低溫跨到高溫都有活性，這在科學上代表它的『適應容忍力』是強還是弱？"
    },
    "26": {
        "name": "游淨O", "score": "77.5", "wrong": "基礎題：5, 13, 14；精熟題：2, 3, 4, 5, 11",
        "analysis": "<b>1. 基礎題 5【胃液消化目標】：</b>混淆了胃液與唾液的化學分工。胃液呈強酸性，主要負責「蛋白質」的初步消化，因此胃消化不良時，首要減食的是蛋白質類食物（如豆腐、肉類）。<br/><br/>"
                    "<b>2. 基礎題 14【自營與異營本質】：</b>對捕蟲植物的營養策略有先入為主的迷思。捕蟲植物依然靠光合作用自營製造養分，與動物完全異營的能量獲取途徑差異最大。<br/><br/>"
                    "<b>3. 精熟題 2【大腸機能迷思】：</b>落入「消化道都能消化」的陷阱。大腸不具有分泌消化液的功能，只負責吸收水分與形成糞便。<br/><br/>"
                    "<b>4. 精熟題 3【酵素溫控實驗與碘液檢測】：</b>這裡有雙重推論陷阱。在 40度C（接近體溫）時唾液澱粉酶活性最強，澱粉被「最快完全分解殆盡」。因為澱粉已經消失，此時滴加碘液，當然會最先不出現藍黑色（呈碘液原本的黃褐色）。<br/><br/>"
                    "<b>5. 精熟題 4、5、11【綜合生理判讀】：</b>對植物光合作用產物（氧氣與呼吸作用對應）、消化道三大養分先後順序（澱粉 -> 蛋白質 -> 脂質），以及原核生物（藍綠菌無葉綠體）的概念缺乏整合。",
        "mindset": "【進階應用與實驗推論鏈條卡關】淨O在精熟題失分較多，顯示面對長題幹與多步驟實驗推論時，容易在邏輯轉換的過程中迷失方向。在「酵素分解速度 -> 澱粉剩餘量 -> 試劑呈色」因果鏈上，容易把「活性大」誤解為「顏色會變深藍黑」，需要徹底重建化學試劑檢測的底層邏輯。",
        "advice": "淨O，這次考卷後面的進階實驗題是不是覺得像在繞口令？別氣餒！這不是你不會，而是我們需要幫大腦建立『推論流程圖』。請你靜下心來，跟著老師的邏輯鏈一步步推導：(1) 胃液裡的胃酸就像是專門剪開『蛋白質』的剪刀。當你胃痛脹氣時，你覺得應該暫時少吃點肉和豆腐（蛋白質），還是少吃蔬菜？(2) 這是一個超級實驗挑戰：如果 40度C 的口水酵素裝成超級大胃王，把試管裡的『澱粉』在三秒鐘內全部吃光光，這時候你滴入專門測澱粉的『碘液』，試管還會變藍黑色嗎？（提示：澱粉都已經不在囉！）」"
    },
    "1": {
        "name": "王佑O", "score": "89.5", "wrong": "基礎題：13；精熟題：3, 8, 11",
        "analysis": "<b>1. 基礎題 13【變形蟲攝食】：</b>對單細胞生物的適應構造較生疏，變形蟲是依靠「吞噬作用」形成食泡來把食物受取進細胞內。<br/><br/>"
                    "<b>2. 精熟題 3【溫控與澱粉酶檢測】：</b>混淆了「酵素活性最高」與「試劑呈色結果」的因果關係。在 40度C 下唾液澱粉酶活性最大，澱粉消耗速度最快，因此滴碘液時會最先不再出現藍黑色。<br/><br/>"
                    "<b>3. 精熟題 8【溫度容忍度曲線】：</b>圖表曲線的橫軸跨度較大時，代表該酵素在溫度變化下依然能保持活性，即對溫度的容忍程度較高。<br/><br/>"
                    "<b>4. 精熟題 11【原核生物胞器界線】：</b>落入葉綠體與光合作用的直覺等號陷阱。藍綠菌是原核生物，沒有被膜包圍的葉綠體（僅有葉綠素）；葉綠體應在真核生物「矽藻」中被發現。",
        "mindset": "【單一觀念強，圖表與多步驟實驗卡關】佑O對基礎名詞的掌握極好（89.5分），是絕對的實力派！目前在面對高溫/低溫影響酵素活性、再轉化為碘液變色與否的「雙重變因推論」時，腦中的思考路徑容易在第二步驟轉向。",
        "advice": "佑O！考出 89.5 分真的很棒，只差 0.5 分就邁入 90 分大關了，老師為你感到驕傲！你這次錯的幾題全部都是『實驗題與圖表推理』。來，把這個實驗推論邏輯刻進你的大腦裡：(1) 在 40度C 體溫環境下，唾液酵素活力全開，把澱粉直接『清空』。既然澱粉編號已經被清光了，碘液滴下去還找不到澱粉，顏色會變成藍黑色，還是維持原本的黃褐色呢？(2) 記住：細菌和藍綠菌是地球上最原始的『原核生物』。只要聽到『原核』兩個字，請立刻在心裡打一個大叉：它們絕對沒有核膜，也絕對沒有『葉綠體』這種高級室內房間！"
    },
    "16": {
        "name": "林沄O", "score": "95.0", "wrong": "精熟題：3, 8 (基礎 20 題全對！)",
        "analysis": "<b>1. 精熟題 3【溫控與澱粉測定實驗】：</b>考生清楚 0度C 活性低、100度C 酵素變性失活，但遇到 40度C（最適溫度）時，容易在最終「碘液呈色結果」推論反轉。正因為 40度C 酵素將澱粉最快水解完畢，所以滴入碘液會最先驗不出藍黑色（呈黃褐色）。<br/><br/>"
                    "<b>2. 精熟題 8【曲線頂點與容忍度辨析】：</b>將曲線的「最高點（最佳溫度）」與「底幅寬度（容忍度）」在閱讀視覺上混淆。圖表中乙曲線在橫軸延伸比甲更寬，代表溫度波動時乙酵素不易失活，對溫度的容忍程度較高。",
        "mindset": "【頂尖卓越，僅在溫控物理量判讀微失手】沄O斬獲 95 分的全班極其優異成績，基礎記憶與理解無懈可擊！唯一的盲點在於「酵素作用越順利 -> 受質（澱粉）越少 -> 檢測劑越不會呈現正反應顏色」的反向剔除邏輯。",
        "advice": "沄O！95 分的高分真的太耀眼了，基礎題 20 題全對，代表你的觀念無比扎實，是超級厲害的生物高手！你這次僅錯的兩題都圍繞在『酵素的溫度實驗』。老師送你兩個頂尖學霸必備的看圖秘訣：(1) 做檢驗實驗時，請記住『反應最快 = 原料（澱粉）消失最快 = 碘液越快不變黑』。千萬不要被『活性強就要顏色深』的心魔給騙了！(2) 下次看酵素圖表時，請分開看兩個重點：往上看到最高點叫作『最佳作用溫度』；往左右看底邊跨越的寬度，寬度越寬就叫作『對溫度的容忍度超強』。掌握這兩點，你就是無敵的 100 分實力！"
    },
    "27": {
        "name": "黃俞O", "score": "88.0", "wrong": "基礎題：4, 13；精熟題：8, 11",
        "analysis": "<b>1. 基礎題 4【纖維素機能】：</b>誤將纖維素當作供能營養素。纖維素無法被人體酵素分解，完全無熱量，攝食目的是為了刺激腸道蠕動。<br/><br/>"
                    "<b>2. 基礎題 13【變形蟲攝食機制】：</b>忽略變形蟲依靠細胞膜變形延伸，進行吞噬作用形成食泡的適應方式。<br/><br/>"
                    "<b>3. 精熟題 8【容忍度圖表解讀】：</b>將曲線高低與橫軸寬窄混為一談。曲線在橫軸跨度較廣者，代表對環境溫度的容忍度較高。<br/><br/>"
                    "<b>4. 精熟題 11【原核與葉綠體迷思】：</b>誤認藍綠菌有葉綠體。藍綠菌屬原核界，缺乏核膜與葉綠體等膜狀胞器；葉綠體應於真核生物矽藻中發現。",
        "mindset": "【典型陷阱全命中】俞O的考卷錯題與諺O完全一模一樣，這代表你們同時落入了出題老師在全卷最精心設計的三大典型盲點陷阱。在生物學名詞的學習上，容易「抓大放小」，記住了大概念（蔬菜很好、光合作用產生糖），卻遺漏了生理限制（人體無纖維素酶、原核生物無膜胞器）。",
        "advice": "俞O，考出 88 分代表你的實力非常不錯喔！不過老師看到你的錯題忍不住會心一笑，因為你跟諺O踩到的陷阱一模一樣，這幾題真的是全班最容易被騙的『大魔王題』！來，幫老師把這三張複習小卡在腦中填滿：(1) 【纖維素小卡】：人體不能消化它，沒有熱量，但吃了可以促進腸道蠕動防便秘。(2) 【藍綠菌小卡】：它是最原始的原核生物，雖然能光合作用，但絕對沒有葉綠體！(3) 【圖表小卡】：看圖表時用手比對橫軸底線，底線長得越寬，代表這隻酵素的溫度容忍度越強！"
    },
    "18": {
        "name": "洪采O", "score": "90.5", "wrong": "基礎題：13, 14；精熟題：11",
        "analysis": "<b>1. 基礎題 13【變形蟲攝食構造】：</b>忽略變形蟲作為單細胞生物，是藉由細胞膜進行吞噬作用形成食泡來補食食物。<br/><br/>"
                    "<b>2. 基礎題 14【自營生產者判斷】：</b>落入「捕蟲植物＝消費者／異營」的直覺陷阱。豬籠草、捕蠅草主要仍然是靠光合作用製造養分（自營生產者），捕蟲是為了生存於貧瘠土壤中補充氮元素，與動物異營生活有本質差異。<br/><br/>"
                    "<b>3. 精熟題 11【生物分類與胞器】：</b>被光合作用生理特徵所誘答。藍綠菌為原核生物，有光合色素但毫無葉綠體；葉綠體應在真核生物（如矽藻）中尋找。",
        "mindset": "【優等學霸，錯題高度集中在生物界分類與營養獲取】采O表現非常優秀，輕鬆突破 90 分大關！你的盲點極度專一，全都圍繞在「自營 vs. 異營」與「原核 vs. 真核」的界線判斷。當生物具有特殊適應（如植物吃蟲、細菌行光合作用）時，容易被特殊習性混淆了它真實的生物學分類界限。",
        "advice": "采O！恭喜拿下超過 90 分的高分，表現實在太出類拔萃了！老師發現你的一個超有趣現象：你錯的三個題目的主題居然完全一模一樣，就是『生物怎麼賺取營養與生物界分類』！這代表你只要解開這個死結，分數直接衝向滿分！請記住這兩個黃金定理：(1) 不管豬籠草吃再多隻蚊子，它的真實職業依然是『靠光合作用自己養自己（自營生物）』！(2) 只要是藍綠菌或細菌，就是原始的『原核界』，它們的細胞裡絕對沒有葉綠體這種高級有膜的胞器！把這兩個分類金鑰握在手裡，你就無敵囉！"
    },
    "14": {
        "name": "賴佑O", "score": "89.0", "wrong": "基礎題：13；精熟題：1, 3, 8",
        "analysis": "<b>1. 基礎題 13【變形蟲攝食】：</b>對於單細胞生物的結構適應不夠熟悉，變形蟲是利用細胞膜進行吞噬作用形成食泡。<br/><br/>"
                    "<b>2. 精熟題 1【消化道酸鹼環境】：</b>忽略人體各消化管腔腔體的酸鹼特化。胃酸（鹽酸）使胃內部維持在強酸性環境，有利於胃蛋白酶作用；而大腸不負責分泌任何消化液。<br/><br/>"
                    "<b>3. 精熟題 3【水浴溫控與呈色反推】：</b>在溫控水浴實驗中，40度C 下唾液酵素活性最強，澱粉被化學水解速度最快。因為試管內的受質（澱粉）迅速消失，滴入碘液會最先不再呈藍黑色。<br/><br/>"
                    "<b>4. 精熟題 8【容忍度圖表解讀】：</b>未能辨別圖表底線橫軸與容忍力的邏輯關係。曲線跨越溫度廣，代表對環境溫度的容忍程度高。",
        "mindset": "【基礎超強，進階實驗與酸鹼環境微卡關】佑O在基礎題高達 95% 正確率，展現了無比堅強的觀念實力（89分）！在精熟題失分主要因為對「消化道各段酸鹼值特異性」與「溫控實驗試劑呈色的結果反推」不夠熟練。",
        "advice": "佑O！基礎題幾乎全對，考出 89 分的超級好成績，代表你的基本功練得超扎實，老師非常肯定你的努力！面對剩下的幾道進階實驗題，我們只需要補上兩個關鍵概念，你就能突破天花板：(1) 記住人體消化道酸鹼值的唯一奇特地圖：『口是中性、胃是強酸（因為有鹽酸）、小腸是鹼性』！這個強酸環境是專門給胃液剪裁蛋白質用的喔。(2) 在酵素實驗裡，那位作用最努力、活性最好的 40度C 酵素，會把試管裡的澱粉全部『吃光光』。既然澱粉都沒了，碘液滴下去當然就不會變藍黑色囉！是不是豁然開朗了呢？"
    }
}

# 3. 生成個人專屬 PDF 報告函數
def generate_pdf_bytes(student_id, student_info):
    buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    doc = SimpleDocTemplate(buffer.name, pagesize=A4, rightMargin=35, leftMargin=35, topMargin=35, bottomMargin=35)
    
    title_style = ParagraphStyle('DocTitle', fontName=active_font, fontSize=17, leading=21, textColor=colors.HexColor("#1A5276"), alignment=1, spaceAfter=12)
    section_style = ParagraphStyle('SectionHeading', fontName=active_font, fontSize=12.5, leading=17, textColor=colors.HexColor("#1A5276"), spaceBefore=12, spaceAfter=6)
    body_style = ParagraphStyle('BodyText', fontName=active_font, fontSize=11, leading=17, textColor=colors.HexColor("#2C3E50"), spaceBefore=3, spaceAfter=4)
    advice_style = ParagraphStyle('AdviceText', fontName=active_font, fontSize=11, leading=17, textColor=colors.HexColor("#0E6251"), spaceBefore=3, spaceAfter=3)
    
    story = []
    story.append(Paragraph(f"<b>《生物科第2回複習考》學生個別化診斷報告</b>", title_style))
    story.append(Spacer(1, 3))
    
    header_data = [
        [Paragraph(f"<b>班級：</b>9年01班 (A組)", body_style), Paragraph(f"<b>座號：</b>{student_id} 號", body_style)],
        [Paragraph(f"<b>姓名：</b>{student_info['name']}", body_style), Paragraph(f"<b>測驗得分：<font color='#E74C3C'><b>{student_info['score']} 分</b></font></b>", body_style)],
        [Paragraph(f"<b>錯題分布：</b><font color='#C0392B'>{student_info['wrong']}</font>", body_style), ""]
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
    
    story.append(Paragraph("<b>一、 誘答機制與核心考點拆解</b>", section_style))
    story.append(Paragraph(student_info['analysis'], body_style))
    
    story.append(Paragraph("<b>二、 思維路徑與學習認知診斷</b>", section_style))
    story.append(Paragraph(student_info['mindset'], body_style))
    story.append(Spacer(1, 6))
    
    advice_content = [
        [Paragraph("<b>三、 老師給你的暖心引導與挑戰：</b>", ParagraphStyle('GreenHeader', fontName=active_font, fontSize=12, leading=16, textColor=colors.HexColor("#0E6251"), spaceAfter=4))],
        [Paragraph(f"「{student_info['advice']}」", advice_style)]
    ]
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
    
    doc.build(story)
    
    with open(buffer.name, "rb") as f:
        pdf_bytes = f.read()
    os.unlink(buffer.name)
    return pdf_bytes


# 6. Streamlit UI 介面設計
st.title("🔬 國中生物科 AI 診斷與群發系統 (A組混班標準模組)")
st.markdown("這個系統由 **Streamlit 驅動**，告別生硬的程式碼。老師只需要『點選、上傳、點擊』即可完成全自動群發。")

# 側邊欄設定
st.sidebar.header("🔐 寄信端安全驗證")
sender_email = st.sidebar.text_input("老師的 Gmail 信箱", placeholder="example@gmail.com")
app_password = st.sidebar.text_input("Google 16位數應用程式密碼", type="password", placeholder="xxxx xxxx xxxx xxxx")

st.sidebar.markdown("---")
test_mode = st.sidebar.checkbox("🛡️ 安全測試模式 (只寄給老師自己)", value=True)
st.sidebar.info("勾選測試模式時，所有發射出去的信箱都會強制導向老師自己的信箱，方便您最後確認。確認無誤再取消勾選正式群發。")

# 主頁面
st.header("📋 數據輸入端")
uploaded_file = st.file_uploader("請上傳您的 A組名單 Excel 檔案 (`901a_st帳號.xlsx`)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # 自動讀取 Excel 內所有的分頁
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names
        
        # 讓老師用下拉選單挑分頁，預設找 "901"
        default_index = sheet_names.index("901") if "901" in sheet_names else 0
        selected_sheet = st.selectbox("🎯 請選擇要發送的班級分頁 (工作表)", sheet_names, index=default_index)
        
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        df.columns = df.columns.str.strip()
        
        if '班別' in df.columns:
            df.rename(columns={'班別': '班級'}, inplace=True)
            
        st.success(f"📋 成功載入名單！偵測到分頁內共有 {len(df)} 筆學生資料。")
        
        # 預覽名單資料
        st.dataframe(df[['班級', '座號', '姓名', 'Google帳號(完整)']].head(10))
        
        # 發射按鈕
        st.markdown("---")
        st.subheader("🚀 派發端")
        if st.button("🔥 開始全自動一鍵群發"):
            if not sender_email or not app_password:
                st.error("❌ 請先在左側邊欄填入您的 Gmail 與 16位數應用程式密碼！")
            else:
                # 測試郵件連線
                try:
                    server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
                    server.login(sender_email, app_password.replace(" ", ""))
                except Exception as login_err:
                    st.error(f"❌ 郵件伺服器驗證失敗，請檢查應用程式密碼是否輸入正確。錯誤: {login_err}")
                    st.stop()
                
                success, skipped, fail = 0, 0, 0
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                total_rows = len(df)
                
                for idx, row in df.iterrows():
                    # 更新進度條
                    progress_bar.progress((idx + 1) / total_rows)
                    
                    student_class = str(row.get('班級', '')).strip()
                    seat_no = str(row.get('座號', '')).strip()
                    student_name = str(row.get('姓名', '')).strip()
                    target_email = str(row.get('Google帳號(完整)', '')).strip()
                    
                    # 智慧型過濾：檢查這個座號有沒有做好的 AI 診斷
                    if seat_no not in STUDENTS_DB:
                        skipped += 1
                        continue
                        
                    student_info = STUDENTS_DB[seat_no]
                    actual_recipient = sender_email if test_mode else target_email
                    
                    status_text.text(f"✍️ 正在處理：901班 {seat_no}號 {student_name}...")
                    
                    try:
                        # 1. 動態在雲端用手寫體生成 PDF 檔案
                        pdf_bytes = generate_pdf_bytes(seat_no, student_info)
                        
                        # 2. 封裝郵件
                        msg = MIMEMultipart()
                        msg['From'] = f"生物科 AI 診斷助教 Pro哥 <{sender_email}>"
                        msg['To'] = actual_recipient
                        msg['Subject'] = f"【個人專屬診斷報告】A組 901班 {seat_no}號 {student_name} - 生物科第2回複習考"
                        
                        mail_body = f"""
                        <html>
                        <body style="font-family: Arial, sans-serif; color: #333; line-height: 1.6;">
                            <div style="max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; padding: 20px; background-color: #fcfcfc;">
                                <h2 style="color: #1A5276; border-bottom: 2px solid #3498DB; padding-bottom: 10px;">🔬 國中生物科：個別化錯題認知診斷報告</h2>
                                <p><b>{student_name} 同學</b>，你好：</p>
                                <p>這份是生物老師與 AI 診斷助教特別為你量身打造的<b>《生物科第2回複習考》深度診斷報告</b>！</p>
                                <p>我們不是只報分數，而是針對你這次考卷中的錯題，深度拆解了出題老師的<b>誘答陷阱</b>，並分析了你學習上的<b>思維路徑與盲點</b>。</p>
                                <div style="background-color: #EBF5FB; border-left: 4px solid #3498DB; padding: 12px; margin: 15px 0;">
                                    💡 <b>老師的小叮嚀：</b><br/>
                                    請下載附件的 PDF 檔案，仔細閱讀「<b>三、 老師給你的暖心引導與挑戰</b>」區域，跟著問題動動腦，把錯過的概念徹底搞懂，下次會考一定能衝出最高分！
                                </div>
                                <p>加油！身為 A 組優秀的同學，期許你精益求精！對報告有任何問題，都歡迎隨時去找生物老師討論喔！</p>
                                <hr style="border: none; border-top: 1px solid #eee; margin: 20px 0;" />
                                <p style="font-size: 12px; color: #888; text-align: center;">此郵件由生物科 AI 系統自動發送，請勿直接回覆此信件。</p>
                            </div>
                        </body>
                        </html>
                        """
                        msg.attach(MIMEText(mail_body, 'html', 'utf-8'))
                        
                        # 3. 夾帶附件
                        attach_file = MIMEApplication(pdf_bytes, _subtype="pdf")
                        attach_file.add_header('Content-Disposition', 'attachment', filename=f"A組_901班_{seat_no}號_{student_name}_生物診斷報告.pdf")
                        msg.attach(attach_file)
                        
                        # 4. 發射
                        server.send_message(msg)
                        success += 1
                        time.sleep(1.2) # 安全防封鎖防刷秒
                    except Exception as send_err:
                        st.warning(f"⚠️ 發送給 {student_name} 時發生問題: {send_err}")
                        fail += 1
                        
                server.quit()
                status_text.empty()
                
                # 報告大功告成成果
                st.balloons()
                st.success("🎉 全自動發信流水線圓滿完工！")
                st.info(f"📊 執行統計：成功發送 {success} 封 | 自動跳過非目標生 {skipped} 封 | 失敗 {fail} 封。")
                if test_mode:
                    st.warning("🛡️ 目前為【安全測試模式】，所有專屬考卷都已經寄到老師您的個人信箱囉！檢查滿意後，取消勾選側邊欄即可正式發全班！")
                    
    except Exception as e:
        st.error(f"❌ 檔案處理發生錯誤，請確認上傳的 Excel 格式是否正確。原因: {e}")
else:
    st.info("💡 請在上方上傳您的學生名單 Excel 檔案，即可開啟自動化操作介面。")
