# ระบบสอบแพทย์พร้อมลงเวลา

ซอร์ส Google Apps Script สำหรับระบบสอบแพทย์ พร้อมบันทึกเวลาเข้า–ออกและสรุปค่าจ้างรายสัปดาห์

## ไฟล์

- `Code.gs` — Apps Script backend และการเชื่อมต่อ Google Sheets/Drive
- `Index.html` — หน้าเว็บระบบสอบ ลงเวลา และสรุปเงินเดือน
- `index.html` — ทางเข้า Cloudflare Pages ซึ่ง redirect ไปยัง Apps Script Web app

## ติดตั้ง Apps Script

1. เปิดโปรเจกต์ Google Apps Script ที่ผูกกับ Google Sheet ของระบบ
2. คัดลอก `Code.gs` และ `Index.html` เข้าโปรเจกต์ โดยตั้งชื่อ HTML file ว่า `Index`
3. ใน Project Settings > Script Properties ตั้ง `INITIAL_PASSWORD` เป็นรหัสผ่านเริ่มต้นที่คาดเดายาก
4. รัน `setup()` จาก Apps Script Editor และอนุญาตสิทธิ์ Google Sheets/Drive ตามที่ระบบร้องขอ
5. Deploy เป็น Web app และกำหนดสิทธิ์เข้าถึงให้เหมาะกับองค์กร

## Cloudflare Pages

Pages entry file `index.html` redirects visitors to the current Apps Script Web app URL. The app itself stays on Apps Script so `google.script.run`, Google Sheets and the existing login continue to work. Deploy the `cloudflare` folder as the Pages output directory, or use the repository root after copying its `index.html` there.

## การตั้งค่าและข้อมูลลับ

- ห้ามเก็บรหัสผ่าน, Discord Bot Token, Webhook URL หรือข้อมูลผู้ใช้จริงในซอร์สโค้ด
- เก็บค่าเริ่มต้นและข้อมูลเชื่อมต่อไว้ใน Script Properties หรือชีตตั้งค่าที่จำกัดสิทธิ์
- รูปหลักฐานลงเวลาเก็บในโฟลเดอร์ Google Drive ส่วนตัวของเจ้าของ Apps Script
- GitHub repository นี้เป็น public ตามที่เจ้าของระบบเลือก อย่าใส่ข้อมูลจากชีตจริงหรือ secret ลงในไฟล์ที่ commit

## ระบบลงเวลา/เงินเดือน

- ใช้บัญชี Username/Password ที่ล็อกอินอยู่ในเว็บเป็นตัวระบุผู้ลงเวลา
- เวลาเข้า–ออกยึดเวลาฝั่งเซิร์ฟเวอร์และต้องแนบรูปหลักฐาน
- รายงานแยกรายวันและรายบุคคล; รอบสัปดาห์วันเสาร์ถึงวันศุกร์
- เกณฑ์ขั้นต่ำ ชั่วโมงทำงาน ช่วงพัก อัตรารายชั่วโมง และโบนัสตั้งค่าได้จากหน้าสรุปเงินเดือน
