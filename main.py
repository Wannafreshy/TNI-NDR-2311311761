import streamlit as st
import pandas as pd
df = pd.read_excel(r"C:\Users\boony\Documents\TNI-NDR-2311311761\TTB-SET-23May2025-6M.xlsx", sheet_name="TTB", skiprows=1)
df.columns = [
"วันที่", "ราคาเปิด", "ราคาสูงสุด", "ราคาต่ำสุด", "ราคาเฉลี่ย", "ราคาปิด",
"เปลี่ยนแปลง", "เปลี่ยนแปลง(%)", "ปริมาณ(พันหุ้น)", "มูลค่า(ล้านบาท)",
"SET Index", "SET เปลี่ยนแปลง(%)"
]
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

# ตั้งค่าเว็บ
st.set_page_config(page_title="วิเคราะห์หุ้น TTB", layout="wide")

st.title("วิเคราะห์แนวโน้มราคาหุ้น TTB (ย้อนหลัง 6 เดือน)")

st.markdown("""
บริษัท **ทีทีบี (TTB)** เป็นธนาคารพาณิชย์ที่มีบทบาทสำคัญในระบบการเงินของประเทศไทย  
ราคาหุ้นของ TTB มีความเคลื่อนไหวตามปัจจัยภายในขององค์กรและสภาวะเศรษฐกิจโดยรวม

ในระบบนี้ คุณสามารถดูข้อมูลย้อนหลัง 6 เดือน เพื่อ:
- วิเคราะห์พฤติกรรมราคาปิดรายวัน
- ประเมินความสัมพันธ์กับดัชนี **SET**
- ตรวจสอบแนวโน้มโดยรวมผ่านโมเดลเชิงเส้น

> *ข้อมูลนี้เหมาะสำหรับนักลงทุนที่ต้องการวิเคราะห์เชิงเทคนิคเบื้องต้น*
""")


# กำหนดพาธไฟล์ Excel
excel_path = r"C:\Users\boony\Documents\TNI-NDR-2311311761\TTB-SET-23May2025-6M.xlsx"

try:
    # อ่านข้อมูล
    df = pd.read_excel(excel_path, sheet_name="TTB", skiprows=1)
    df.columns = [
        "วันที่", "ราคาเปิด", "ราคาสูงสุด", "ราคาต่ำสุด", "ราคาเฉลี่ย", "ราคาปิด",
        "เปลี่ยนแปลง", "เปลี่ยนแปลง(%)", "ปริมาณ(พันหุ้น)", "มูลค่า(ล้านบาท)",
        "SET Index", "SET เปลี่ยนแปลง(%)"
    ]

    # ทำความสะอาดข้อมูล
    df = df[~df["วันที่"].isna() & ~df["วันที่"].str.contains("วันที่")]
    df["วันที่"] = df["วันที่"].apply(convert_thai_date)
    df["วันที่"] = pd.to_datetime(df["วันที่"])
    df = df.dropna().sort_values("วันที่")

    # ตารางข้อมูลย้อนหลัง 6 เดือน
    st.subheader("ตารางข้อมูลย้อนหลัง 6 เดือน")
    st.dataframe(df.style.format({
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
        "SET เปลี่ยนแปลง(%)": "{:.2f}%"
    }).set_table_styles([
        {'selector': 'thead th', 'props': [('background-color', '#1b4f72'), ('color', 'white'), ('font-weight', 'bold')]},
        {'selector': 'tbody tr:nth-child(even)', 'props': [('background-color', '#f2f9ff')]},
        {'selector': 'tbody tr:hover', 'props': [('background-color', '#b3d7ff')]}
    ]))

except Exception as e:
    st.error(f"เกิดข้อผิดพลาด: {e}")
# df = df[~df["วันที่"].isna() & ~df["วันที่"].str.contains("วันที่")]
# df["วันที่"] = df["วันที่"].apply(convert_thai_date)
# df["วันที่"] = pd.to_datetime(df["วันที่"])
# df = df.dropna()
# df.head(5)
# print(df["ราคาปิด"].describe())
# print(df[df["ราคาปิด"] == df["ราคาปิด"].max()])
# print(df[["ราคาปิด", "SET Index"]].corr())
# from sklearn.linear_model import LinearRegression
# import numpy as np
# import matplotlib.pyplot as plt
# import matplotlib
# matplotlib.rcParams['font.family'] = 'DejaVu Sans'
# df_sorted = df.sort_values("วันที่")
# X = df_sorted["วันที่"].map(pd.Timestamp.toordinal).values.reshape(-1, 1)
# y = df_sorted["ราคาปิด"].values
# model = LinearRegression()
# model.fit(X, y)
# trend = model.predict(X)
# plt.figure(figsize=(12, 6))
# plt.plot(df_sorted["วันที่"], y, label="Actual Closing Price")
# plt.plot(df_sorted["วันที่"], trend, label="Trend (Linear Regression)",
# linestyle="--", color="red")
# plt.title("TTB Closing Price Trend")
# plt.xlabel("Date")
# plt.ylabel("Closing Price (Baht)")
# plt.legend()
# plt.grid(True)
# plt.tight_layout()
# plt.show()