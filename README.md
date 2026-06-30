# Aloha Party — CMMBA 34

ระบบลงทะเบียน + บริจาคสด สำหรับงาน **CMMBA 34 Aloha Party**
สร้างด้วย **Google Apps Script** (HTML Service + Google Sheets + Drive OCR)

## ไฟล์ในโปรเจกต์

| ไฟล์ | หน้าที่ |
|------|---------|
| `รหัส.gs` | Backend — ลงทะเบียน, OCR สลิป, บันทึกบริจาค, Top Spenders |
| `Index.html` | หน้าฟอร์มลงทะเบียน (ผู้ใช้กรอก) |
| `Display.html` | หน้าจอ Live Donations (โปรเจกเตอร์) — เปิดด้วย `?p=display` |

## การติดตั้ง

1. เปิด [Google Apps Script](https://script.google.com) สร้างโปรเจกต์ใหม่
2. คัดลอกโค้ดจากแต่ละไฟล์ไปวางให้ตรงชื่อ
3. **เปิด Drive API**: Services (➕ ในแถบซ้าย) → Drive API → Add
4. Deploy → New deployment → Web app
5. ตั้งค่า Sheet/Folder ID ใน `CONFIG` ที่หัวไฟล์ `รหัส.gs`

## ลิงก์

- หน้าลงทะเบียน: `<WEB_APP_URL>`
- หน้าจอ Live: `<WEB_APP_URL>?p=display`
