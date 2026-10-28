# SUMMARY — แยก MF เป็น SUM-EDI.pdf + SUMMARY.pdf

โปรแกรมอ่านไฟล์ MF (CARGO MANIFEST ทั้งลำ แบบ .xls/.xlsx ที่มีช่อง `1) SHED NO.` / `2) ประเทศ (XX)` / `3) CARGO MOVEMENT`)
แล้วแยก B/L เป็นกลุ่ม LCH, LKB (GATE2) และสินค้า Transit แต่ละประเทศ จากนั้นทำ 2 ไฟล์

- `SUM-EDI.pdf` ใบสรุป A4 หัวจดหมาย HEUNG A LINE: จำนวนตู้ 20'/40', จำนวน B/L, น้ำหนักรวม (KGS), จำนวนหีบห่อ ของแต่ละกลุ่ม และ TOTAL
- `SUMMARY.pdf` หน้า "SUMMARY OF INWARD CARGO MANIFEST DISCHARGING AT LAEM CHABANG" กลุ่มละ 1 หน้า (แนวนอน) พร้อมลายเซ็น

## เงื่อนไขการแยกกลุ่ม

| กลุ่ม | B/L ที่อยู่ในกลุ่ม | หน้า SUMMARY |
|---|---|---|
| LCH | B/L ที่ไม่เข้ากลุ่มอื่น | LAEM CHABANG |
| LKB (GATE2) | `1) SHED NO. 0332` | LAEM CHABANG / LKB (ESCO GATE 2) |
| MYANMAR / LAOS / MALAYSIA | `3) CARGO MOVEMENT (7-TRANSIT)` และรหัสประเทศ `MM` / `LA` / `MY` | \*CARGO INTRANSIT TO … P.D.R. VIA LAEM CHABANG |

- จำนวน B/L นับรวม House B/L (…A, …B)
- ขนาดตู้จากรหัสประเภทตู้ของ B/L: `22G0` = 20' · `45G0` / `45R1` = 40' (ตู้เดียวกันนับครั้งเดียว)
- ยอดรวมทุกกลุ่มต้องเท่ากับ **GRAND TOTAL** ของ MF (โปรแกรมแจ้งว่าตรงหรือไม่ตรง)
- ลำดับหน้า SUMMARY: LCH → MYANMAR → LAOS → MALAYSIA → LKB

## ใช้งานผ่านเว็บ (ไม่ต้องติดตั้งอะไร)

เปิดที่ **https://rattanao.github.io/Sum-edi/**

1. ลากไฟล์ MF (.xls / .xlsx) มาวาง เว็บจะแยกกลุ่มและแสดงตารางทันที (คลิกชื่อกลุ่มเพื่อดูรายการ B/L)
2. ช่อง **ลายเซ็น** เลือกรูปลายเซ็น (.png พื้นโปร่งใส หรือ .jpg) ครั้งเดียว เบราว์เซอร์จะจำไว้ในเครื่องนั้น กด **ลบลายเซ็น** เพื่อเอาออก
3. ดูตัวอย่างทุกหน้าด้านล่าง แล้วกด **ดาวน์โหลด SUM-EDI.pdf + SUMMARY.pdf**

ไฟล์ถูกประมวลผลในเบราว์เซอร์ของเครื่องที่ใช้อยู่เท่านั้น ไม่มีการอัปโหลดข้อมูลหรือลายเซ็นไปที่ใด ตัวเว็บคือไฟล์ [`index.html`](index.html)

## ใช้งานด้วยโปรแกรม Python (Windows)

สิ่งที่ต้องมี: Python 3.9 ขึ้นไป และฟอนต์ Times New Roman / Tahoma ของ Windows

```bash
pip install -r requirements.txt
```

1. วางไฟล์ MF ไว้ในโฟลเดอร์ `input`
2. (ไม่บังคับ) วางรูปลายเซ็นชื่อ `assets/signature.png` (ไฟล์นี้ไม่ถูกส่งขึ้น GitHub)
3. ดับเบิลคลิก `run-summary.bat` หรือสั่ง

   ```bash
   python -X utf8 scripts/sum_mf.py .
   ```

ผลลัพธ์อยู่ที่ `output/SUM-EDI.pdf` และ `output/SUMMARY.pdf`
ระบุไฟล์หรือโฟลเดอร์เองได้: `python -X utf8 scripts/sum_mf.py "D:\งาน\MF.xls" "D:\งาน\output"`

## ใช้กับ Claude Code

โฟลเดอร์นี้เป็น skill ชื่อ `summary` (ไฟล์ [`SKILL.md`](SKILL.md)) พิมพ์ `summary` หรือ "ทำ summary จาก MF" แล้วแนบไฟล์ MF

---
Project by Ms. Rattana Mudlae · Import DOC. · HEUNG A LINE (THAILAND) CO.,LTD.
