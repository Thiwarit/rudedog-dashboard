# Rudedog Stock Dashboard

Dashboard สำหรับจัดการสต็อกสินค้า Rudedog

## ฟีเจอร์หลัก
- แสดงข้อมูลสต็อก SM และ FG
- คำนวณมูลค่าสินค้าอัตโนมัติ
- กราฟแสดงยอดขายและแนวโน้ม
- ระบบเตือนสินค้าที่ไม่มีราคา

## วิธีใช้งาน
1. อัปโหลดไฟล์ Excel ที่มี 6 ชีต:
   - Sheet1 (ข้อมูลหลัก)
   - ราคา
   - SM
   - ตัดออก
   - FG รุ่นทำตลาด
   - SM ใช้งาน

2. ดูผลลัพธ์ใน Dashboard

## การติดตั้งสำหรับใช้งานภายใน
```bash
pip install streamlit pandas openpyxl
streamlit run dashboard.py
```

---
สร้างโดย Claude AI สำหรับ Rudedog