Project ENG-494
# วิเคราะห์แนวโน้มราคาหุ้น TTB ด้วย Streamlit
โปรเจกต์นี้เป็นเว็บแอปที่สร้างด้วย [Streamlit](https://streamlit.io) สำหรับวิเคราะห์แนวโน้มราคาหุ้น **TTB (TMBThanachart Bank Public Company Limited)** โดยใช้ข้อมูลย้อนหลัง 6 เดือนจากไฟล์ Excel  
แสดงผลในรูปแบบอินเทอร์แอคทีฟ พร้อมตารางราคาหุ้นที่ตกแต่งด้วย CSS และกราฟแสดงเส้นแนวโน้ม (Trend Line) ด้วย Linear Regression

---
คุณสมบัติ (Features)
- โหลดและแปลงข้อมูลจาก Excel ที่มีวันที่แบบไทย
- แสดงตารางราคาหุ้นย้อนหลัง พร้อมการจัดรูปแบบตัวเลขและสี
- คำนวณสถิติเชิงพรรณนา เช่น ค่าเฉลี่ย สูงสุด ต่ำสุด
- คำนวณและแสดงความสัมพันธ์ (Correlation) ระหว่างราคาปิดหุ้นกับดัชนี SET Index
- แสดงกราฟราคาปิดพร้อมเส้นแนวโน้ม (Trend Line) ด้วย Linear Regression
- ตกแต่ง UI ด้วย CSS ให้ดูอ่านง่าย
- รองรับภาษาไทย
---
## ไลบรารีที่ใช้ (Dependencies)

> ไฟล์ `requirements.txt` :

```text
streamlit
pandas
matplotlib
scikit-learn
openpyxl 
```
## Python version
```text
Python 3.11.9
```
##CANVA

https://www.canva.com/design/DAGoXpNQucw/EkKTKHP90F8xF5j8rSUcQg/edit?utm_content=DAGoXpNQucw&utm_campaign=designshare&utm_medium=link2&utm_source=sharebutton

