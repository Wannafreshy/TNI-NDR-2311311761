try:
    df = pd.read_excel(excel_path, sheet_name="TTB", skiprows=1)
    df.columns = [
        "วันที่", "ราคาเปิด", "ราคาสูงสุด", "ราคาต่ำสุด", "ราคาเฉลี่ย", "ราคาปิด",
        "เปลี่ยนแปลง", "เปลี่ยนแปลง(%)", "ปริมาณ(พันหุ้น)", "มูลค่า(ล้านบาท)",
        "SET Index", "SET เปลี่ยนแปลง(%)"
    ]
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

    # จากตรงนี้เป็นการใช้งานข้อมูล
    st.subheader("ตารางข้อมูลย้อนหลัง 6 เดือน")
    st.dataframe(style_table(df))

    # สถิติราคาปิด
    st.subheader("สถิติราคาปิดหุ้น TTB")
    st.write(df["ราคาปิด"].describe().apply(lambda x: f"{x:.2f}"))

    max_price = df["ราคาปิด"].max()
    max_price_date = df.loc[df["ราคาปิด"] == max_price, "วันที่"].dt.strftime('%d %b %Y').values[0]
    st.markdown(f"**ราคาปิดสูงสุดในช่วงนี้:** {max_price:.2f} บาท เมื่อวันที่ {max_price_date}")

    corr = df[["ราคาปิด", "SET Index"]].corr().iloc[0, 1]
    st.markdown(f""" ... ข้อมูลความสัมพันธ์ ... """, unsafe_allow_html=True)


    fig, ax = plt.subplots(figsize=(8, 6), dpi=100)
    ax.plot(df["วันที่"], df["ราคาปิด"], marker='o', linestyle='-', color='#2874a6', label='ราคาปิด')

    # Linear Regression
    X = df["วันที่"].map(pd.Timestamp.toordinal).values.reshape(-1, 1)
    y = df["ราคาปิด"].values
    model = LinearRegression()
    model.fit(X, y)
    trend = model.predict(X)
    ax.plot(df["วันที่"], trend, color='#d35400', linestyle='--', label='เส้นแนวโน้ม')

    ax.set_xlabel("วันที่", fontsize=12)
    ax.set_ylabel("ราคาปิด (บาท)", fontsize=12)
    ax.set_title("กราฟราคาปิดหุ้น TTB กับเส้นแนวโน้ม", fontsize=16)
    ax.tick_params(axis='x', labelrotation=45)
    ax.grid(True, linestyle='--', alpha=0.5)
    ax.legend(fontsize=12)
    fig.tight_layout()
    fig.subplots_adjust(left=0.15, right=0.95, top=0.9, bottom=0.2)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.pyplot(fig, use_container_width=False)

    # สรุปผล
    trend_direction = "เพิ่มขึ้น" if trend[-1] > trend[0] else "ลดลง" if trend[-1] < trend[0] else "คงที่"
    corr_val = corr

    st.markdown(f""" ... สรุปแนวโน้ม ... """, unsafe_allow_html=True)

except Exception as e:
    st.error(f"เกิดข้อผิดพลาด: {e}")
