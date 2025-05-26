import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.linear_model import LinearRegression
import matplotlib

thai_font = 'Tahoma'
matplotlib.rcParams['font.family'] = thai_font
# พาธไฟล์ Excel
excel_path = "TTB-SET-23May2025-6M.xlsx"

try:
    df = pd.read_excel(excel_path, sheet_name="TTB", skiprows=1)
    df.columns = [
        "วันที่", "ราคาเปิด", "ราคาสูงสุด", "ราคาต่ำสุด", "ราคาเฉลี่ย", "ราคาปิด",
        "เปลี่ยนแปลง", "เปลี่ยนแปลง(%)", "ปริมาณ(พันหุ้น)", "มูลค่า(ล้านบาท)",
        "SET Index", "SET เปลี่ยนแปลง(%)"
    ]
    
# แปลงวันที่ไทยเป็น YYYY-MM-DD
    thai_months = {
        "ม.ค.": "01", "ก.พ.": "02", "มี.ค.": "03", "เม.ย.": "04",
        "พ.ค.": "05", "มิ.ย.": "06", "ก.ค.": "07", "ส.ค.": "08",
        "ก.ย.": "09", "ต.ค.": "10", "พ.ย.": "11", "ธ.ค.": "12"
    }

    def convert_thai_date(thai_date_str):
        for th, num in thai_months.items():
            if th in thai_date_str:
                day, month_th, year_th = thai_date_str.replace(",", "").split()
                month = thai_months[month_th]
                year = int(year_th) - 543
                return f"{year}-{month}-{int(day):02d}"
        return None

    df = df[~df["วันที่"].isna() & ~df["วันที่"].str.contains("วันที่")]
    df["วันที่"] = df["วันที่"].apply(convert_thai_date)
    df["วันที่"] = pd.to_datetime(df["วันที่"])
    df = df.dropna().sort_values("วันที่")

    # ฟังก์ชันตกแต่งตาราง
    def style_table(df):
        def color_change(val):
            if isinstance(val, str):
                return ''
            if val > 0:
                return 'color: #27ae60; font-weight: 600;'  
            elif val < 0:
                return 'color: #c0392b; font-weight: 600;'  
            else:
                return ''

        styled_df = (
            df.style
            .format({
                "ราคาเปิด": "{:,.2f}",
                "ราคาสูงสุด": "{:,.2f}",
                "ราคาต่ำสุด": "{:,.2f}",
                "ราคาเฉลี่ย": "{:,.2f}",
                "ราคาปิด": "{:,.2f}",
                "เปลี่ยนแปลง": "{:,.2f}",
                "เปลี่ยนแปลง(%)": "{:.2f}%",
                "ปริมาณ(พันหุ้น)": "{:,.0f}",
                "มูลค่า(ล้านบาท)": "{:,.0f}",
                "SET Index": "{:,.2f}",
                "SET เปลี่ยนแปลง(%)": "{:.2f}%",
            })
            .set_table_styles([
            {'selector': 'thead th', 'props': [
                ('color', 'white'),
                ('font-weight', '700'),
                ('text-align', 'center'),
                ('padding', '10px'),
            ]},          
        ])
        .applymap(color_change, subset=["เปลี่ยนแปลง(%)", "SET เปลี่ยนแปลง(%)"])
        .set_properties(
            subset=["ราคาเปิด", "ราคาสูงสุด", "ราคาต่ำสุด", "ราคาเฉลี่ย", "ราคาปิด",
                    "เปลี่ยนแปลง", "เปลี่ยนแปลง(%)", "ปริมาณ(พันหุ้น)", "มูลค่า(ล้านบาท)",
                    "SET Index", "SET เปลี่ยนแปลง(%)"], 
            **{'font-family': "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"}
        )
    )
        return styled_df 

# ตั้งฟอนต์ภาษาไทย matplotlib
# thai_font = 'Tahoma'
# matplotlib.rcParams['font.family'] = thai_font

# ตั้งค่าเพจ
    st.set_page_config(page_title="วิเคราะห์หุ้น TTB", layout="wide")

# CSS 
    st.markdown("""
<style>
.stApp {
    background: linear-gradient(to bottom right,#F5F5DC, #E0FFFF) !important;
    color: #212529 !important;
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif !important;
    font-size: 16px !important;
}
h1, h2 {
    color: #1a5276 !important;
    font-weight: 800 !important;
}
h3 {
    color: #2471a3 !important;
    font-weight: 700 !important;
}
a {
    color: #1f618d !important;
    text-decoration: none !important;
}
a:hover {
    color: #154360 !important;
    text-decoration: underline !important;
}
table {
    border-collapse: collapse !important;
    width: 100% !important;
    border: 1px solid #dee2e6 !important;
}
thead th {
    background-color: #2e86c1 !important;
    color: white !important;
    padding: 10px !important;
    text-align: center !important;
    font-weight: bold !important;
}
tbody td {
    padding: 10px !important;
    border-bottom: 1px solid #dee2e6 !important;
    text-align: right !important;
}
tbody tr:nth-child(even) {
    background-color: #f4faff !important;
}
tbody tr:hover {
    background-color: #d6eaf8 !important;
}
.stButton>button, .stDownloadButton>button {
    background-color: #3498db !important;
    color: white !important;
    border: none !important;
    padding: 0.5em 1em !important;
    border-radius: 6px !important;
    font-size: 16px !important;
    font-weight: 600 !important;
    transition: background-color 0.3s !important;
}
.stButton>button:hover, .stDownloadButton>button:hover {
    background-color: #2e86c1 !important;
    color: white !important;
}
footer {
    visibility: hidden !important;
}
</style>
""", unsafe_allow_html=True)

# หัวข้อ
    st.markdown("""
<h1 style="font-weight: 800; color:#1a5276; margin-bottom: 0;">
    วิเคราะห์แนวโน้มราคาหุ้น TTB
</h1>
<p style="font-size: 1.45rem; color:#2471a3; margin-top: 2px;">
    (ราคาย้อนหลัง 6 เดือน : ข้อมูลล่าสุด ณ 23 พ.ค. 2568)
</p>
""", unsafe_allow_html=True)
    st.markdown("""
บริษัท ทีทีบี (TTB) เป็นหนึ่งในธนาคารชั้นนำของประเทศไทย ราคาหุ้นของบริษัทมีความผันผวนตามสถานการณ์ทางเศรษฐกิจและการเงิน  
ในเว็บนี้ เราจะดูข้อมูลย้อนหลัง 6 เดือน โดยทำความสะอาดข้อมูลและวิเคราะห์แนวโน้มราคาปิด พร้อมกราฟแสดงราคาหุ้นและเส้นเทรนด์  
""")



    st.subheader("ตารางข้อมูลย้อนหลัง 6 เดือน")
    st.dataframe(style_table(df))

    # สถิติราคาปิด
    st.subheader("สถิติราคาปิดหุ้น TTB")
    st.write(df["ราคาปิด"].describe().apply(lambda x: f"{x:.2f}"))

    max_price = df["ราคาปิด"].max()
    max_price_date = df.loc[df["ราคาปิด"] == max_price, "วันที่"].dt.strftime('%d %b %Y').values[0]
    st.markdown(f"**ราคาปิดสูงสุดในช่วงนี้:** {max_price:.2f} บาท เมื่อวันที่ {max_price_date}")

    corr = df[["ราคาปิด", "SET Index"]].corr().iloc[0,1]
    st.markdown(f"""
   <div style="
    background-color: #eaf2f8; 
    padding: 15px 20px; 
    border-left: 5px solid #2874a6; 
    border-radius: 8px;
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    color: #222222;
    line-height: 1.5;
    max-width: 600px;
">
    <p style="font-weight: 700; font-size: 18px; margin-bottom: 8px;">
        ความสัมพันธ์ระหว่างราคาปิดกับ SET Index: <span style="color:#d35400;">{corr:.4f}</span>
    </p>
    <p style="font-size: 16px; margin-bottom: 0;">
        ความสัมพันธ์นี้บอกเราว่า ราคาหุ้น <strong>TTB</strong> มีแนวโน้มเคลื่อนไหวในทิศทางเดียวกับดัชนี SET Index ค่อนข้างสูง
    </p>
</div>
""", unsafe_allow_html=True)

    # กราฟราคาปิดและเทรนด์
    st.subheader("กราฟราคาปิดและเส้นแนวโน้ม (Trend Line)")

    fig, ax = plt.subplots(figsize=(8, 6), dpi=100)
    ax.plot(df["วันที่"], df["ราคาปิด"], marker='o', linestyle='-', color='#2874a6', label='ราคาปิด')

# Linear Regression
    X = df["วันที่"].map(pd.Timestamp.toordinal).values.reshape(-1, 1)
    y = df["ราคาปิด"].values
    model = LinearRegression()
    model.fit(X, y)
    trend = model.predict(X)
    ax.plot(df["วันที่"], trend, color='#d35400', linestyle='--', label='เส้นแนวโน้ม')

# ปรับ Label และ Title
    ax.set_xlabel("วันที่", fontsize=12)
    ax.set_ylabel("ราคาปิด (บาท)", fontsize=12)
    ax.set_title("กราฟราคาปิดหุ้น TTB กับเส้นแนวโน้ม", fontsize=16)
    ax.tick_params(axis='x', labelrotation=45)
    ax.grid(True, linestyle='--', alpha=0.5)
    ax.legend(fontsize=12)

# ปรับ layout ไม่ให้กราฟเบียด
    fig.tight_layout()
# หรือแบบละเอียด:
    fig.subplots_adjust(left=0.15, right=0.95, top=0.9, bottom=0.2)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
     st.pyplot(fig, use_container_width=False)


 # สรุปผล
    trend_direction = "เพิ่มขึ้น" if trend[-1] > trend[0] else "ลดลง" if trend[-1] < trend[0] else "คงที่"
    corr_val = corr

    st.markdown(f"""
    <div style="
    background-color: #f0f8ff; 
    padding: 20px; 
    border-radius: 12px; 
    border: 2px solid #87ceeb;
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    color: #222222;
    line-height: 1.6;
    max-width: 700px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.08);
    margin-bottom: 30px;
">
    <p style="font-size: 18px; margin-bottom: 12px;">
        ราคาหุ้น <strong style="color: #2874a6;">TTB</strong> มีความ
        <span style="color: #d35400; font-weight: bold;">ผันผวน</span> ในช่วง 6 เดือนที่ผ่านมา
    </p>
    <p style="font-size: 18px; margin-bottom: 12px;">
        แนวโน้มโดยรวมเป็นไปในทิศทาง  
        <span style="color: #27ae60; font-weight: bold;">{trend_direction}</span>
    </p>
    <p style="font-size: 18px; margin-bottom: 12px;">
        ราคาปิดมีความสัมพันธ์กับดัชนี SET ในระดับ  
        <span style="color: {'#c0392b' if corr_val < 0 else '#2980b9'}; font-weight: bold;">
            {corr_val:.2f}
        </span>  
        แสดงถึงการตอบสนองต่อภาวะตลาดโดยรวม
    </p>
    <p style="font-size: 16px; color: #555555;">
        นักลงทุนควรติดตามข่าวสารและสถานการณ์เศรษฐกิจเพิ่มเติมประกอบการตัดสินใจลงทุน
    </p>
</div>
""", unsafe_allow_html=True)


except Exception as e:
    st.error(f"เกิดข้อผิดพลาด: {e}") 
