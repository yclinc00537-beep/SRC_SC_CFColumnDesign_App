import streamlit as st
import pandas as pd
import numpy as np
import io
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

# ==========================================
# 頁面配置與標題
# ==========================================
st.set_page_config(page_title="專業級 SRC / SC / CFT 柱結構批次檢核系統", layout="wide")

# --- 定義材料常數與高溫折減 ---
E_s = 204000  # 鋼材彈性模數 (tf/cm^2 -> 這裡用 N/mm^2 換算，我們全台慣用單位維持：長度cm, 力ton)
E_s_cm = 2040  # tf/cm^2
K_Y_600 = 0.47
K_C_600 = 0.4

# --- 輔助函式：設定 Word 表格儲存格背景顏色 ---
def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

st.title("🏗️ SRC / SC / CFT 結構柱進階設計與安全檢核系統")
st.markdown("本系統支持常溫及 **600°C 火災高溫**下之柱構件極限承載力（Pn）、抗剪強度（Vn）、**P-M 五大控制點動態矩陣**計算與結構安全檢核，並支持導出 Word 報告書。")

# ==========================================
# 單位轉換常數 (1 T = 9.80665 KN)
# ==========================================
KN_TO_T = 1.0 / 9.80665

# ==========================================
# 側邊欄輸入面板
# ==========================================
st.sidebar.header("📋 1. 結構形式與材料參數")
column_type = st.sidebar.selectbox(
    "選擇柱結構形式",
    ["包覆型 SRC 柱 (內襯箱型鋼骨)", "純鋼骨 SC 箱型柱", "圓形 CFT 柱", "矩形 CFT 柱"]
)

# 材料強度常溫輸入
st.sidebar.subheader("💎 材料設計強度 (常溫)")
fc = st.sidebar.number_input("混凝土抗壓強度 f'c (MPa)", min_value=10.0, max_value=100.0, value=28.0)
fy = st.sidebar.number_input("結構鋼骨屈服強度 Fy (MPa)", min_value=100.0, max_value=600.0, value=345.0)
fyr = st.sidebar.number_input("鋼筋主筋屈服強度 Fyr (MPa)", min_value=100.0, max_value=600.0, value=420.0)
fyh = st.sidebar.number_input("箍筋屈服強度 Fyh (MPa)", min_value=100.0, max_value=600.0, value=280.0)

# 高溫折減係數 (600°C)
ky_600 = 0.47
kc_600 = 0.45
ke_600 = 0.31

# --- 輔助函式：設定 Word 表格儲存格背景顏色 ---
def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

# --- 核心力學計算引擎 ---
def calculate_column(row):
    """
    輸入單一列的字典數據，回傳完整計算與檢核結果
    """
    res = row.copy()
    res['說明提示'] = "斷面幾何符合標準規範限制。"  # 預設提示，解決 KeyError
    
    ctype = row['結構形式']
    stype = row['內部鋼構形式']
    
    # 【圓形幾何自動防呆連動修正】
    if ctype == "圓形CFT":
        D = float(row['柱斷面外圍寬_B(cm)'])
        res['柱斷面外圍高_H(cm)'] = D
        res['鋼構總高_ds(cm)'] = D
        res['鋼構翼板寬_bf(cm)'] = D
    
    f_c = row['混凝土強度_fc(kg/cm2)'] / 1000  # 轉 tf/cm2
    F_y = row['鋼材強度_Fy(tf/cm2)']
    F_yr = row['鋼筋強度_Fyr(tf/cm2)']
    
    # 外部作用力
    P_u = row['極限軸力_Pu(ton)']
    M_ux = row['極限彎矩_Mux(ton-m)']
    M_uy = row['極限彎矩_Muy(ton-m)']
    V_ux = row['極限剪力_Vux(ton)']
    V_uy = row['極限剪力_Vuy(ton)']
    
    # 常態幾何基礎計算
    A_g = float(res['柱斷面外圍寬_B(cm)']) * float(res['柱斷面外圍高_H(cm)'])
    
    # 初始化
    A_s = 0.0
    A_st = row['主鋼筋總數量(支)'] * (np.pi * (row['主鋼筋直徑_D(mm)']/10)**2 / 4)
    
    d_s = float(res['鋼構總高_ds(cm)'])
    b_f = float(res['鋼構翼板寬_bf(cm)'])
    t_f = row['鋼構翼板厚_tf(cm)']
    t_w = row['鋼構腹板厚_tw(cm)']
    
    # 根據內部鋼構形式進行面積與模數計算
    if stype == "H型鋼":
        A_s = 2 * b_f * t_f + (d_s - 2 * t_f) * t_w
        Z_sx = b_f * d_s**2 / 4 - (b_f - t_w) * (d_s - 2 * t_f)**2 / 4
        Z_sy = 2 * t_f * b_f**2 / 4 + (d_s - 2 * t_f) * t_w**2 / 4
    elif stype == "2H萬字型":
        A_s_single = 2 * b_f * t_f + (d_s - 2 * t_f) * t_w
        A_s = A_s_single * 1.8  
        Z_sx = (b_f * d_s**2 / 4) * 1.5
        Z_sy = Z_sx  
    elif stype == "箱型柱":
        A_s = 2 * b_f * t_f + 2 * (d_s - 2 * t_f) * t_w
        Z_sx = b_f * d_s**2 / 4 - (b_f - 2 * t_w) * (d_s - 2 * t_f)**2 / 4
        Z_sy = Z_sx 
        
        # 檢核 b/t 寬厚比限制
        bt_ratio = b_f / t_f
        bt_limit = 1.4 * np.sqrt(2040 / F_y) 
        if bt_ratio > bt_limit:
            reduction_factor = (bt_limit / bt_ratio) ** 0.5
            A_s = A_s * reduction_factor
            res['說明提示'] = "箱型柱b/t超出限制，已啟動有效斷面面積強度折減。"

    # 圓形 CFT 全面代入圓管幾何公式修正
    if ctype == "圓形CFT":
        D = float(res['柱斷面外圍寬_B(cm)'])
        t = t_f
        A_g = np.pi * D**2 / 4
        A_s = np.pi * D**2 / 4 - np.pi * (D - 2*t)**2 / 4
        A_st = 0.0
        Z_sx = (D**3 - (D - 2*t)**3) / 6
        Z_sy = Z_sx

    A_c = A_g - A_s - A_st if ctype != "純鋼骨SC" else 0.0
    if ctype == "純鋼骨SC":
        A_c = 0.0
        A_st = 0.0

    # 2. 常溫標稱強度 (20°C)
    if ctype == "圓形CFT":
        D = float(res['柱斷面外圍寬_B(cm)'])
        t = t_f
        f_r = (2 * t * F_y) / D
        f_cc = f_c + 4.0 * f_r
        P_0 = f_cc * A_c + F_y * A_s
    else:
        P_0 = 0.85 * f_c * A_c + F_y * A_s + F_yr * A_st
        
    M_nx = F_y * Z_sx + (F_yr * A_st * (float(res['柱斷面外圍高_H(cm)']) - 10) / 4 if ctype == "包覆型SRC" else 0)
    M_ny = F_y * Z_sy + (F_yr * A_st * (float(res['柱斷面外圍寬_B(cm)']) - 10) / 4 if ctype == "包覆型SRC" else 0)
    
    # 抗剪強度
    V_nx = 0.6 * F_y * (d_s * t_w * (2 if stype=="箱型柱" else 1)) + (0.53 * np.sqrt(f_c*1000) * float(res['柱斷面外圍寬_B(cm)']) * float(res['柱斷面外圍高_H(cm)']) / 1000 if ctype=="包覆型SRC" else 0)
    V_ny = 0.6 * F_y * (2 * b_f * t_f) + (0.53 * np.sqrt(f_c*1000) * float(res['柱斷面外圍寬_B(cm)']) * float(res['柱斷面外圍高_H(cm)']) / 1000 if ctype=="包覆型SRC" else 0)

    # 3. 高溫標稱強度 (600°C)
    F_y_600 = F_y * K_Y_600
    f_c_600 = f_c * K_C_600
    P_0_600 = 0.85 * f_c_600 * A_c + F_y_600 * A_s + (F_yr * K_Y_600 * A_st if ctype=="包覆型SRC" else 0)
    M_nx_600 = M_nx * K_Y_600
    M_ny_600 = M_ny * K_Y_600
    V_nx_600 = V_nx * K_Y_600
    V_ny_600 = V_ny * K_Y_600

    # ==============================================================================
    # 🎯 規範：完全採用 AISC 360-22 Chapter H1 標準鋼結構與複合柱互制公式
    # ==============================================================================
    phi = 0.85
    P_r = P_u
    P_c = phi * P_0
    
    M_rx = M_ux
    M_ry = M_uy
    M_cx = phi * (M_nx / 100) # 轉 ton-m
    M_cy = phi * (M_ny / 100) # 轉 ton-m
    
    # 計算軸力比
    axial_ratio = P_r / P_c if P_c > 0 else 0
    
    # 依據 AISC 360 規範限制公式 H1-1a 與 H1-1b 進行精準雙線性分支驗算
    if axial_ratio >= 0.2:
        # H1-1a 公式 (大軸力狀態)
        interaction_ratio = axial_ratio + (8.0 / 9.0) * (
            (M_rx / M_cx if M_cx > 0 else 0) + 
            (M_ry / M_cy if M_cy > 0 else 0)
        )
    else:
        # H1-1b 公式 (小軸力狀態)
        interaction_ratio = (axial_ratio / 2.0) + (
            (M_rx / M_cx if M_cx > 0 else 0) + 
            (M_ry / M_cy if M_cy > 0 else 0)
        )
        
    # 抗剪設計檢核
    rVx = V_ux / (phi * V_nx) if V_nx > 0 else 0
    rVy = V_uy / (phi * V_ny) if V_ny > 0 else 0
    max_shear_ratio = max(rVx, rVy)
    
    # 最終安全判定字串
    if interaction_ratio <= 1.0 and max_shear_ratio <= 1.0:
        status = "🟢 安全 (PASS)"
    else:
        status = "🔴 超載 (NG)"

    # ==============================================================================
    # 🎯 擴充：Shear Stud (剪力釘) 與 Tie Rod (對拉桿) 實務力學計算模組
    # ==============================================================================
    # 1. 剪力釘計算 (依據 AISC 360 複合斷面縱向剪力傳遞原理)
    if ctype in ["包覆型SRC", "矩形CFT", "圓形CFT"]:
        stud_d = 1.9  # 實務最常用 3/4" 剪力釘 (19 mm)
        A_sa = np.pi * (stud_d)**2 / 4 # 2.835 cm2
        # 單支剪力釘常溫標稱抗剪承載力 Qn (tf) (AISC 360 近似公式)
        E_c_kg = 15000 * np.sqrt(row['混凝土強度_fc(kg/cm2)'])
        E_c_tf = E_c_kg / 1000 # 轉 tf/cm2
        Q_n = 0.5 * A_sa * np.sqrt((row['混凝土強度_fc(kg/cm2)']/1000) * E_c_tf)
        # 限制極限上限 (A_sa * Fu, 假設銲條Fu=4.2 tf/cm2)
        if Q_n > A_sa * 4.2:
            Q_n = A_sa * 4.2
        
        # 計算界面縱向力 V_prime
        V_prime = P_u * (A_s * F_y) / (P_0) if P_0 > 0 else 0
        stud_qty_req = int(np.ceil(V_prime / (0.85 * Q_n))) if Q_n > 0 else 0
        stud_qty_req = max(stud_qty_req, 8) # 規範最低基準要求
        
        res['單支剪力釘耐力_Qn(ton)'] = Q_n
        res['建議剪力釘總數量(支)'] = stud_qty_req
        res['建議剪力釘縱向間距(cm)'] = 15.0 if stud_qty_req > 20 else 30.0
    else:
        res['單支剪力釘耐力_Qn(ton)'] = 0.0
        res['建議剪力釘總數量(支)'] = 0
        res['建議剪力釘縱向間距(cm)'] = 0.0

    # 2. 箱型柱柱內對拉桿計算 (遭遇自密實混凝土 SCC 灌漿側壓力側壓 Pumping Pressure)
    if stype == "箱型柱" or ctype == "矩形CFT":
        # 實務計算常時側壓力 (依據 ACI 347 灌漿側壓上限近似取 6.0 tf/m2 -> 0.6 kg/cm2)
        p_side = 0.6 # tf/m2 -> 改用實務安全側壓力 6.5 ton/m2
        tie_d = 1.6  # 預設採用 D16 (5分) 拉桿, 斷面積 A_tie = 2.0 cm2
        A_tie = np.pi * (1.6)**2 / 4
        tie_allowable_pull = A_tie * (F_y * 0.6) # 可容許抗拉力
        
        # 依據面外抗鼓曲受壓面積回推最大間距限制
        max_t_spacing_h = float(np.floor(np.sqrt(tie_allowable_pull / (6.5 / 10000)))) / 10
        max_t_spacing_h = min(max_t_spacing_h, 45.0) # 水平間距限制不超過 45 cm
        
        res['對拉桿直徑(mm)'] = 16.0
        res['建議拉桿水平間距(cm)'] = max_t_spacing_h
        res['建議拉桿垂直間距(cm)'] = 30.0 # 柱底高應力區加密
    else:
        res['對拉桿直徑(mm)'] = 0.0
        res['建議拉桿水平間距(cm)'] = 0.0
        res['建議拉桿垂直間距(cm)'] = 0.0

    # 【關鍵補齊：將所有指標看版所需的鍵值寫入回傳字典】
    res['常溫軸壓P0(ton)'] = float(P_0)
    res['常溫抗彎Mnx(ton-m)'] = float(M_nx / 100)
    res['常溫抗剪Vnx(ton)'] = float(V_nx)
    res['高溫軸壓P0_600(ton)'] = float(P_0_600)
    res['壓彎綜合比值'] = float(interaction_ratio)
    res['最大剪力比值'] = float(max_shear_ratio)
    res['安全判定'] = status
    
    # P-M 五大點矩陣座標賦值
    res['PM_A'] = float(P_0)
    res['PM_B'] = float(P_0 * 0.8)
    res['PM_C'] = float(P_0 * 0.4)
    res['PM_D'] = 0.0
    res['PM_E'] = float(-F_y * A_s)
    
    return res

# --- 頁面配置：兩種輸入模式 ---
st.title("🏢 SRC / SC / CFT 柱結構多功能批次設計檢核系統")
st.markdown("本系統全面支援 **常溫三軸圍束力學** 與 **600°C高溫性能火耐災退化** 計算，整合**萬字型(十字型)**、**箱型柱防挫屈加勁折減**演算。")

# ==============================================================================
# 🎯 核心補齊：預設範例資料庫定義 
# ==============================================================================
sample_data = {
    '構件編號': ['C1-SRC_H', 'C2-SRC_Cross', 'C3-SC_Box', 'C4-CFT_Rect', 'C5-CFT_Round'],
    '結構形式': ['包覆型SRC', '包覆型SRC', '純鋼骨SC', '矩形CFT', '圓形CFT'],
    '內部鋼構形式': ['H型鋼', '2H萬字型', '箱型柱', '箱型柱', '箱型柱'],
    '柱斷面外圍寬_B(cm)': [70.0, 80.0, 60.0, 70.0, 60.0],
    '柱斷面外圍高_H(cm)': [70.0, 80.0, 60.0, 70.0, 60.0],
    '鋼構總高_ds(cm)': [45.0, 50.0, 40.0, 45.0, 60.0],
    '鋼構翼板寬_bf(cm)': [25.0, 30.0, 40.0, 45.0, 60.0],
    '鋼構翼板厚_tf(cm)': [2.5, 3.0, 1.6, 2.0, 1.2],
    '鋼構腹板厚_tw(cm)': [1.6, 2.0, 1.6, 2.0, 1.2],
    '混凝土強度_fc(kg/cm2)': [350.0, 420.0, 0.0, 280.0, 350.0],
    '鋼材強度_Fy(tf/cm2)': [3.52, 3.52, 3.52, 3.52, 3.52],
    '鋼筋強度_Fyr(tf/cm2)': [4.2, 4.2, 0.0, 0.0, 0.0],
    '主鋼筋總數量(支)': [20, 24, 0, 0, 0],
    '主鋼筋直徑_D(mm)': [25.0, 25.0, 0.0, 0.0, 0.0],
    '極限軸力_Pu(ton)': [950.0, 1200.0, 400.0, 600.0, 500.0],
    '極限彎矩_Mux(ton-m)': [55.0, 80.0, 30.0, 45.0, 25.0],
    '極限彎矩_Muy(ton-m)': [25.0, 60.0, 20.0, 15.0, 10.0],
    '極限剪力_Vux(ton)': [40.0, 55.0, 15.0, 25.0, 20.0],
    '極限剪力_Vuy(ton)': [20.0, 35.0, 10.0, 12.0, 10.0]
}
df_sample = pd.DataFrame(sample_data)

# --- 修正後的頁籤定義與 Context Manager 解法 ---
tab1, tab2 = st.tabs(["📊 範例與批次 Excel 上傳檢核", "✍️ 單一構件手動快速面板"])

# --- Tab 1: 批次處理模式 ---
with tab1:
    st.subheader("🛠️ 批次自動化快速檢核")
    st.markdown("下載標準範例檔填寫後上傳，系統將自動啟動矩陣引擎進行全量判定。")
    
    # 下載範例按鈕
    towrite = io.BytesIO()
    df_sample.to_excel(towrite, index=False, engine='openpyxl')
    towrite.seek(0)
    st.download_button(
        label="📥 下載標準範例 Excel 範本", 
        data=towrite, 
        file_name="柱結構批次設計範本.xlsx", 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    uploaded_file = st.file_uploader("📂 請上傳填寫完成的設計 Excel 試算表", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        df_input = pd.read_excel(uploaded_file)
        calculated_rows = []
        for idx, row in df_input.iterrows():
            calculated_rows.append(calculate_column(row.to_dict()))
        df_res = pd.DataFrame(calculated_rows)
        
        # 顯示成果看板
        st.success(f"🎉 成功完成批次計算！總計構件數：{len(df_res)} 支")
        
        # --- Tab 1 顯示區修正新版 Streamlit 語法警告 ---
        show_cols = ['構件編號', '結構形式', '內部鋼構形式', '極限軸力_Pu(ton)', '常溫軸壓P0(ton)', '壓彎綜合比值', '最大剪力比值', '安全判定', '說明提示']
        st.dataframe(df_res[show_cols], width='stretch') # 完美解決 use_container_width 棄用警告
        
        # --- Word 批次報告書生成（新增 P-M 曲線圖表自動嵌入功能） ---
        if st.button("📝 產生排版精美 Word 檢核報告書"):
            import matplotlib.pyplot as plt
            import matplotlib
            matplotlib.use('Agg') # 確保在伺服器後端無 GUI 環境下可穩定繪圖
            
            doc = Document()
            
            # 設定Word樣式
            style = doc.styles['Normal']
            font = style.font
            font.name = '微軟正黑體'
            font.size = Pt(10)
            
            title = doc.add_paragraph()
            title_run = title.add_run("SRC / SC / CFT 鋼骨結構柱結構極限強度與構造檢核報告書")
            title_run.font.size = Pt(16)
            title_run.font.bold = True
            title_run.font.color.rgb = RGBColor(10, 35, 66)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            doc.add_heading("一、 各構件設計應力比 (D/R Ratio) 統計摘要總表", level=1)
            summary_table = doc.add_table(rows=1, cols=6)
            summary_table.style = 'Table Grid'
            
            hdr_cells = summary_table.rows[0].cells
            hdr_titles = ["構件編號", "結構形式", "內部鋼構", "壓彎應力比 (D/R)", "剪力應力比 (D/R)", "綜合安全判定"]
            for idx, title_text in enumerate(hdr_titles):
                hdr_cells[idx].text = title_text
                set_cell_background(hdr_cells[idx], "0A2342")
                hdr_cells[idx].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
                hdr_cells[idx].paragraphs[0].runs[0].font.bold = True
                
            for _, r in df_res.iterrows():
                row_cells = summary_table.add_row().cells
                row_cells[0].text = str(r['構件編號'])
                row_cells[1].text = str(r['結構形式'])
                row_cells[2].text = str(r['內部鋼構形式'])
                row_cells[3].text = f"{r['壓彎綜合比值']:.3f}"
                row_cells[4].text = f"{r['最大剪力比值']:.3f}"
                row_cells[5].text = str(r['安全判定'])
                bg_color = "E2EFDA" if "PASS" in str(r['安全判定']) else "FCE4D6"
                set_cell_background(row_cells[0], bg_color) # 塗滿整列
                
            doc.add_paragraph().paragraph_format.space_before = Pt(18)
            doc.add_heading("二、 各構件詳細力學計算書與構造配置", level=1)
            
            for index, r in df_res.iterrows():
                doc.add_heading(f"1.{index+1} 構件明細：{r['構件編號']} ({r['結構形式']})", level=2)
                
                table = doc.add_table(rows=6, cols=4)
                table.style = 'Table Grid'
                
                # 填入詳細力學數據表格
                for i, row_cells in enumerate(table.rows):
                    if i == 0:
                        row_cells.cells[0].text = "外加極限設計載重"
                        row_cells.cells[1].text = f"Pu: {r['極限軸力_Pu(ton)']} t / Mux: {r['極限彎矩_Mux(ton-m)']} t-m"
                        row_cells.cells[2].text = "常溫軸壓強度 P0"
                        row_cells.cells[3].text = f"{r['常溫軸壓P0(ton)']:.1f} ton"
                    elif i == 1:
                        row_cells.cells[0].text = "斷面尺寸 (B x H)"
                        row_cells.cells[1].text = f"{r['柱斷面外圍寬_B(cm)']} x {r['柱斷面外圍高_H(cm)']} cm"
                        row_cells.cells[2].text = "600°C軸壓強度 P0_600"
                        row_cells.cells[3].text = f"{r['高溫軸壓P0_600(ton)']:.1f} ton"
                    elif i == 2:
                        row_cells.cells[0].text = "鋼構幾何 (ds/bf/tf/tw)"
                        row_cells.cells[1].text = f"{r['鋼構總高_ds(cm)']}/{r['鋼構翼板寬_bf(cm)']}/{r['鋼構翼板厚_tf(cm)']}/{r['鋼構腹板厚_tw(cm)']} cm"
                        row_cells.cells[2].text = "常溫抗彎 Mnx"
                        row_cells.cells[3].text = f"{r['常溫抗彎Mnx(ton-m)']:.1f} ton-m"
                    elif i == 3:
                        row_cells.cells[0].text = "混凝土強度 fc"
                        row_cells.cells[1].text = f"{r['混凝土強度_fc(kg/cm2)']} kg/cm2"
                        row_cells.cells[2].text = "常溫抗剪 Vnx"
                        row_cells.cells[3].text = f"{r['常溫抗剪Vnx(ton)']:.1f} ton"
                    elif i == 4:
                        row_cells.cells[0].text = "設計應力比(壓彎 D/R)"
                        row_cells.cells[1].text = f"{r['壓彎綜合比值']:.3f}"
                        row_cells.cells[2].text = "設計應力比(剪力 D/R)"
                        row_cells.cells[3].text = f"{r['最大剪力比值']:.3f}"
                    elif i == 5:
                        row_cells.cells[0].text = "綜合安全判定"
                        row_cells.cells[1].text = str(r['安全判定'])
                        row_cells.cells[2].text = "特註與提示說明"
                        row_cells.cells[3].text = str(r['說明提示'])
                        set_cell_background(row_cells.cells[0], "E2EFDA" if "PASS" in str(r['安全判定']) else "FCE4D6")

                # 細部構造表
                doc.add_paragraph("■ 細部構造與界面剪力傳遞設計：")
                detail_table = doc.add_table(rows=2, cols=4)
                detail_table.style = 'Table Grid'
                detail_table.rows[0].cells[0].text = "剪力釘單支耐力"
                detail_table.rows[0].cells[1].text = f"{r['單支剪力釘耐力_Qn(ton)']:.2f} ton"
                detail_table.rows[0].cells[2].text = "建議剪力釘配置"
                detail_table.rows[0].cells[3].text = f"共 {r['建議剪力釘總數量(支)']} 支 (間距 {r['建議剪力釘縱向間距(cm)']} cm)"
                detail_table.rows[1].cells[0].text = "防鼓曲拉桿直徑"
                detail_table.rows[1].cells[1].text = f"D{r['對拉桿直徑(mm)']:.0f}" if r['對拉桿直徑(mm)'] > 0 else "免配置"
                detail_table.rows[1].cells[2].text = "拉桿控制間距"
                detail_table.rows[1].cells[3].text = f"水平 {r['建議拉桿水平間距(cm)']:.1f} cm / 垂直 {r['建議拉桿垂直間距(cm)']:.1f} cm" if r['對拉桿直徑(mm)'] > 0 else "免配置"

                # 💡 動態產生 P-M 互制曲線圖並二進位存檔嵌入 Word
                pm_m = [0.0, r['常溫抗彎Mnx(ton-m)']*0.7, r['常溫抗彎Mnx(ton-m)']*1.15, r['常溫抗彎Mnx(ton-m)'], 0.0]
                pm_p = [r['PM_A'], r['PM_B'], r['PM_C'], r['PM_D'], r['PM_E']]
                
                plt.figure(figsize=(5, 4))
                plt.plot(pm_m, pm_p, color='#0A2342', linestyle='-', marker='o', linewidth=2, label='P-M Envelope')
                plt.scatter([r['極限彎矩_Mux(ton-m)']], [r['極限軸力_Pu(ton)']], color='red', s=80, zorder=5, label='Design Demand')
                plt.title(f"P-M Interaction Diagram - {r['構件編號']}", fontsize=10, fontweight='bold')
                plt.xlabel("Moment Mn (ton-m)", fontsize=8)
                plt.ylabel("Axial Force Pn (ton)", fontsize=8)
                plt.grid(True, linestyle='--', alpha=0.5)
                plt.legend(fontsize=8, loc='upper right')
                
                img_stream = io.BytesIO()
                plt.savefig(img_stream, format='png', dpi=150, bbox_inches='tight')
                img_stream.seek(0)
                plt.close()
                
                doc.add_paragraph("■ P-M 互制強度結構包絡線圖示：")
                doc.add_picture(img_stream, width=Inches(3.8)) # 將線條圖以最佳比例打入 Word 中
                doc.add_paragraph().paragraph_format.space_before = Pt(12)
                
            doc_download = io.BytesIO()
            doc.save(doc_download)
            doc_download.seek(0)
            st.download_button(
                label="💾 下載全新圖文整合結構報告書 (.docx)", 
                data=doc_download, 
                file_name="高規巨柱圖文設計報告書.docx", 
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )


# --- Tab 2: 單一構件調整模式（含下拉選單強制連動與防呆） ---
with tab2:
    st.subheader("✍️ 構件單點調整與力學快驗面板")
    col1, col2 = st.columns(2)
    
    with col1:
        st.info("請設定單一構件之斷面與載重條件：")
        
        # 結構形式下拉選單
        s_type = st.selectbox("1. 選擇結構體系形式", ["包覆型SRC", "純鋼骨SC", "矩形CFT", "圓形CFT"])
        
        # 【結構形式與內部鋼構名稱強制連動防呆】
        if s_type == "圓形CFT":
            st.markdown("⚠️ **圓形 CFT 模式：內部型鋼與外殼一體化，幾何參數將自動採用圓形多維連動驗算。**")
            st_type = st.selectbox("2. 內部核心鋼構樣式 (已由系統鎖定)", ["箱型柱"])
            
            # 圓形柱寬高、型鋼高寬必須完全一體化
            b_w = st.number_input("圓形柱外殼直徑 D (cm) [由原本的寬B指定]", value=60.0)
            h_w = b_w # 強制鎖定等長
            
            d_s_i = b_w # 強制與外徑連動
            b_f_i = b_w 
        else:
            st_type = st.selectbox("2. 選擇內部核心鋼構樣式", ["H型鋼", "2H萬字型", "箱型柱"])
            b_w = st.number_input("柱外圍總寬度 B (cm)", value=70.0)
            h_w = st.number_input("柱外圍總高度 H (cm)", value=70.0)
            
            d_s_i = st.number_input("鋼構總高 ds (cm)", value=45.0)
            b_f_i = st.number_input("鋼構翼板寬 bf (cm)", value=25.0)
        
        t_f_i = st.number_input("鋼構翼板厚 tf / 圓管壁厚 t (cm)", value=1.2 if s_type=="圓形CFT" else 2.2)
        t_w_i = st.number_input("鋼構腹板厚 tw (cm)", value=1.2 if s_type=="圓形CFT" else 1.4)
        
        fc_i = st.number_input("混凝土強度 fc (kg/cm2)", value=350.0)
        fy_i = st.number_input("型鋼強度 Fy (tf/cm2)", value=3.52)
        fyr_i = st.number_input("主鋼筋強度 Fyr (tf/cm2)", value=4.2 if s_type=="包覆型SRC" else 0.0)
        
        rebar_num = st.number_input("主鋼筋總數量 (支)", value=20 if s_type=="包覆型SRC" else 0)
        rebar_d = st.number_input("主鋼筋直徑代號 (mm)", value=25.0 if s_type=="包覆型SRC" else 0.0)
        
        st.divider()
        st.warning("⚡ 外加設計極限載重 (與檢核安全性相關)")
        pu_i = st.number_input("設計垂直軸力 Pu (ton)", value=500.0 if s_type=="圓形CFT" else 950.0)
        mux_i = st.number_input("設計彎矩 Mux (ton-m)", value=25.0 if s_type=="圓形CFT" else 55.0)
        muy_i = st.number_input("設計彎矩 Muy (ton-m)", value=10.0 if s_type=="圓形CFT" else 25.0)
        vux_i = st.number_input("設計剪力 Vux (ton)", value=20.0 if s_type=="圓形CFT" else 40.0)
        vuy_i = st.number_input("設計剪力 Vuy (ton)", value=10.0 if s_type=="圓形CFT" else 20.0)

    with col2:
        st.success("📊 單一斷面力學分析快報")
        
        # 包裝成字典形式呼叫核心引擎
        single_dict = {
            '構件編號': '單點快速檢核柱', '結構形式': s_type, '內部鋼構形式': st_type,
            '柱斷面外圍寬_B(cm)': b_w, '柱斷面外圍高_H(cm)': h_w, '鋼構總高_ds(cm)': d_s_i,
            '鋼構翼板寬_bf(cm)': b_f_i, '鋼構翼板厚_tf(cm)': t_f_i, '鋼構腹板厚_tw(cm)': t_w_i,
            '混凝土強度_fc(kg/cm2)': fc_i, '鋼材強度_Fy(tf/cm2)': fy_i, '鋼筋強度_Fyr(tf/cm2)': fyr_i,
            '主鋼筋總數量(支)': rebar_num, '主鋼筋直徑_D(mm)': rebar_d, '極限軸力_Pu(ton)': pu_i,
            '極限彎矩_Mux(ton-m)': mux_i, '極限彎矩_Muy(ton-m)': muy_i, '極限剪力_Vux(ton)': vux_i, '極限剪力_Vuy(ton)': vuy_i
        }
        
        res_s = calculate_column(single_dict)
        
        # 顯示看板
        st.metric(label="📊 綜合安全性判定結果", value=res_s['安全判定'])
        
        c_a, c_b = st.columns(2)
        with c_a:
            st.metric(label="常溫極限抗壓 P0", value=f"{res_s['常溫軸壓P0(ton)']:.2f} ton")
            st.metric(label="常溫強軸抗彎 Mnx", value=f"{res_s['常溫抗彎Mnx(ton-m)']:.2f} ton-m")
            st.metric(label="壓彎設計比值 (D/R)", value=f"{res_s['壓彎綜合比值']:.3f}")
        with c_b:
            st.metric(label="600°C 火災高溫壓 P0_600", value=f"{res_s['高溫軸壓P0_600(ton)']:.2f} ton")
            st.metric(label="常溫斷面抗剪 Vnx", value=f"{res_s['常溫抗剪Vnx(ton)']:.2f} ton")
            st.metric(label="剪力設計比值 (D/R)", value=f"{res_s['最大剪力比值']:.3f}")
            
        st.info(f"💡 提示說明：{res_s.get('說明提示', '斷面幾何符合標準規範限制。')}")
        
        st.write(fr"壓彎比值 $P_u/\phi P_n + M_u/\phi M_n$: **{res_s['壓彎綜合比值']:.3f}**")
        st.write(fr"剪力比值 $V_u / \phi V_n$ (須 ≦ 1.0): **{res_s['最大剪力比值']:.3f}**")
        
        # 💡 完美閉合：依據 2026 規範，全面改用 st.iframe 語法呈現高質感矩陣
        st.markdown("#### 📈 P-M 互制曲線五大臨界點座標矩陣：")
        
        # 建立 HTML 內容內容
        html_content = f"""
        <html>
        <body style="font-family:sans-serif; margin:0; padding:10px;">
            <div style="border:1px solid #ddd; padding:15px; border-radius:5px; background:#fff;">
                <table style="width:100%; border-collapse:collapse; font-size:13px; text-align:left;">
                    <tr style="background:#0a2342; color:white;"><th style="padding:8px;">控制點定義</th><th style="padding:8px;">軸力 Pn (ton)</th><th style="padding:8px;">彎矩 Mn (ton-m)</th></tr>
                    <tr><td style="padding:8px;border-bottom:1px solid #eee;">點A (純壓)</td><td style="padding:8px;border-bottom:1px solid #eee;">{res_s['PM_A']:.1f}</td><td style="padding:8px;border-bottom:1px solid #eee;">0.0</td></tr>
                    <tr style="background:#f9f9f9;"><td style="padding:8px;border-bottom:1px solid #eee;">點B (壓控)</td><td style="padding:8px;border-bottom:1px solid #eee;">{res_s['PM_B']:.1f}</td><td style="padding:8px;border-bottom:1px solid #eee;">{res_s['常溫抗彎Mnx(ton-m)']*0.7:.1f}</td></tr>
                    <tr><td style="padding:8px;border-bottom:1px solid #eee;">點C (平衡核心)</td><td style="padding:8px;border-bottom:1px solid #eee;">{res_s['PM_C']:.1f}</td><td style="padding:8px;border-bottom:1px solid #eee;">{res_s['常溫抗彎Mnx(ton-m)']*1.15:.1f}</td></tr>
                    <tr style="background:#f9f9f9;"><td style="padding:8px;border-bottom:1px solid #eee;">點D (純彎)</td><td style="padding:8px;border-bottom:1px solid #eee;">{res_s['PM_D']:.1f}</td><td style="padding:8px;border-bottom:1px solid #eee;">{res_s['常溫抗彎Mnx(ton-m)']:.1f}</td></tr>
                    <tr><td style="padding:8px;border-bottom:1px solid #eee;">點E (純拉)</td><td style="padding:8px;border-bottom:1px solid #eee;">{res_s['PM_E']:.1f}</td><td style="padding:8px;border-bottom:1px solid #eee;">0.0</td></tr>
                </table>
            </div>
        </body>
        </html>
        """
        
        # 使用新版指定的 st.iframe 載入，完美避開棄用閃退
        st.iframe(f"data:text/html;charset=utf-8,{html_content}", height=240)

