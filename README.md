# SUMMARY — แยก MF เป็น SUM-EDI.pdf + SUMMARY.pdf (LAEM CHABANG / BANGKOK + ตู้ SOC)

โปรแกรมอ่านไฟล์ MF / MANIFEST (CARGO MANIFEST ทั้งลำ แบบ .xls/.xlsx ที่มีช่อง `1) SHED NO.` / `2) ประเทศ (XX)` / `3) CARGO MOVEMENT`)
และไฟล์ CNTRS.xls (CONTAINER LIST, ไม่ใส่ก็ได้) แล้วแยก B/L เป็นกลุ่มตามท่าเรือ / Transit แยกตู้ SOC จากนั้นทำ 2 ไฟล์

- `SUM-EDI.pdf` ใบสรุป A4 หัวจดหมาย HEUNG A LINE: จำนวนตู้ 20'/40', จำนวน B/L, น้ำหนักรวม (KGS), จำนวนหีบห่อ ของแต่ละกลุ่ม, TOTAL, บรรทัดตู้ SOC และ GRAND TOTAL
- `SUMMARY.pdf` หน้า "SUMMARY OF INWARD CARGO MANIFEST DISCHARGING AT LAEM CHABANG / BANGKOK" กลุ่มละ 1 หน้า (แนวนอน) พร้อมลายเซ็น

เนื้อหาในเอกสารเป็นภาษาอังกฤษทั้งหมด ยกเว้นหัวจดหมาย (ที่อยู่ด้านบน)

## เงื่อนไขการแยกกลุ่ม

| กลุ่ม | B/L ที่อยู่ในกลุ่ม | หน้า SUMMARY |
|---|---|---|
| LCH / BANGKOK | B/L ที่ไม่เข้ากลุ่มอื่น (ท่าเรือตาม `PORT OF DISCHARGE` ใน MF) | LAEM CHABANG / BANGKOK |
| LKB (GATE2) | `1) SHED NO. 0332` | LAEM CHABANG / LKB (ESCO GATE 2) |
| MYANMAR / LAOS / MALAYSIA | `3) CARGO MOVEMENT (7-TRANSIT)` และรหัสประเทศ `MM` / `LA` / `MY` | \*CARGO INTRANSIT TO … P.D.R. VIA … |

- **ตู้ SOC** = ตู้ที่อยู่ใน CNTRS แต่ไม่มีใน MF แยกตามเจ้าของ (CONSIGNEE เช่น `BEN LINE AGENCIES (THAILAND) LTD.` → BEN LINE, `NS INTERNATIONAL (THAILAND) CO.,LTD.` → NS) และขนาดตู้ แสดงเป็นบรรทัดแยกใต้ TOTAL (เอกสารไม่มีสี) พร้อม `[SUM-EDI, MANIFEST AS PER ATTACHED]` แล้วรวมเป็น **GRAND TOTAL** (ตัวอย่าง: 60 + 1 + 21 = 82 CNTRS)
- จำนวน B/L นับรวม House B/L (…A, …B)
- ขนาดตู้จาก CNTRS ถ้าไม่มี ใช้รหัสประเภทตู้: `22G0` = 20' · `45G0` / `45R1` = 40' (ตู้เดียวกันนับครั้งเดียว)
- ยอดรวมทุกกลุ่มต้องเท่ากับ **GRAND TOTAL** ของ MF (โปรแกรมแจ้งว่าตรงหรือไม่ตรง)
- ท้าย SUM-EDI: ชื่อ / ตำแหน่ง / เบอร์โทร ของผู้จัดทำ (แก้ได้ เพราะมีพนักงานหลายคน)

## ใช้งานผ่านเว็บ (ไม่ต้องติดตั้งอะไร)

เปิดที่ **https://rattanao.github.io/Sum-edi/**

1. ลากไฟล์ MF และ CNTRS (.xls / .xlsx) มาวางพร้อมกันในช่องเดียว เว็บแยกเองว่าไฟล์ไหนเป็น MF ไฟล์ไหนเป็น CNTRS แล้วแสดงตารางทันที (คลิกชื่อกลุ่มเพื่อดู B/L คลิกแถว SOC เพื่อดูเลขตู้)
2. ช่อง **ชื่อผู้จัดทำ / ตำแหน่ง / เบอร์โทร** แก้เป็นของตัวเอง เบราว์เซอร์จะจำไว้ในเครื่องนั้น
3. ช่อง **ลายเซ็น** เลือกรูปลายเซ็น (.png พื้นโปร่งใส หรือ .jpg) ครั้งเดียว กด **ลบลายเซ็น** เพื่อเอาออก
4. ดูตัวอย่างทุกหน้าด้านล่าง แล้วกด **ดาวน์โหลด SUM-EDI.pdf + SUMMARY.pdf**

ไฟล์ถูกประมวลผลในเบราว์เซอร์ของเครื่องที่ใช้อยู่เท่านั้น ไม่มีการอัปโหลดข้อมูล ชื่อ หรือลายเซ็นไปที่ใด ตัวเว็บคือไฟล์ [`index.html`](index.html)

## ใช้งานด้วยโปรแกรม Python (Windows)

สิ่งที่ต้องมี: Python 3.9 ขึ้นไป และฟอนต์ Times New Roman / Tahoma ของ Windows

```bash
pip install -r requirements.txt
```

1. วางไฟล์ MF (หรือ MANIFEST) และ CNTRS ไว้ในโฟลเดอร์ `input`
2. (ไม่บังคับ) วางรูปลายเซ็นชื่อ `assets/signature.png` (ไฟล์นี้ไม่ถูกส่งขึ้น GitHub)
3. ดับเบิลคลิก `run-summary.bat` หรือสั่ง

   ```bash
   python -X utf8 scripts/sum_mf.py .
   ```

ผลลัพธ์อยู่ที่ `output/SUM-EDI.pdf` และ `output/SUMMARY.pdf`

เปลี่ยนชื่อผู้จัดทำ: `python -X utf8 scripts/sum_mf.py . --name "MISS ..." --position "Import Documentation" --tel "02-059 ...."`
ระบุไฟล์เอง: `python -X utf8 scripts/sum_mf.py "D:\งาน\MANIFEST.xls" "D:\งาน\output" --cntrs "D:\งาน\CNTRS.xls"`

## ใช้กับ Claude Code

โฟลเดอร์นี้เป็น skill ชื่อ `summary` (ไฟล์ [`SKILL.md`](SKILL.md)) พิมพ์ `summary` / `SUM-EDI` หรือ "ทำ summary จาก MF" แล้วแนบไฟล์ MF และ CNTRS

---
Project by Ms. Rattana Mudlae · Import DOC. · HEUNG A LINE (THAILAND) CO.,LTD.
