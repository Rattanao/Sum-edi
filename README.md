# SUM-EDI — แยก MF ตาม DELIVERY เป็น SUM-EDI.pdf + SUMMARY.pdf (+ ตู้ SOC)

โปรแกรมอ่านไฟล์ MF / MANIFEST (CARGO MANIFEST ทั้งลำ เป็น Excel หรือ PDF ก็ได้ ที่มีช่อง `1) SHED NO.` / `2) ประเทศ (XX)` / `3) CARGO MOVEMENT`)
และไฟล์ CNTRS (CONTAINER LIST FOR FEEDER เป็น Excel หรือ PDF ชื่อไฟล์อะไรก็ได้ เช่น CNTRS.xls / INWARD.xls / CNTRS-A0.pdf แยกหลายไฟล์ก็ได้) แล้วแยก B/L ตาม **DELIVERY** แยกตู้ SOC จากนั้นทำ 2 ไฟล์

- `SUM-EDI.pdf` ใบสรุป A4 หัวจดหมาย HEUNG A LINE: AGENT / TAX ID / VESSEL / DISCHARGE และตาราง DELIVERY | TOTAL CNTRS | TOTAL B/L | TOTAL GROSS WEIGHT | TOTAL PKGS, GRAND TOTAL, บรรทัดตู้ SOC และ GRAND TOTAL รวม SOC
- `SUMMARY.pdf` หน้า "SUMMARY OF INWARD CARGO MANIFEST DISCHARGING AT …" (แนวนอน) พร้อมลายเซ็น
- `MF-LCH.pdf` / `MF-LKR.pdf` / `MF-MYANMAR.pdf` … (เมื่อ MF เป็น Excel) MF แยกตาม DELIVERY หน้าตาเหมือนที่พิมพ์จากระบบ มีหน้า SUMMARY ของ DELIVERY นั้นเป็นใบปะหน้า และตัวแดงจาก shed ทุกหน้า (LKR = `LKB (ESCO GATE 2)` ตาม GATE · BMT / SCT = `LAEM CHABANG / BMT`)
- `CNTRS.pdf` (เมื่อ CNTRS เป็น Excel ซึ่งระบบใส่ Port Of Delivery เป็น THLCH ทุก shed) พิมพ์ CONTAINER LIST ใหม่แยกตาม DELIVERY จาก SHED NUMBER (0332 = THLKR มี `/GATE n` สีแดงทุกหน้า) และแยกตู้ SOC แต่ละเจ้าของ

เนื้อหาในเอกสารเป็นภาษาอังกฤษทั้งหมด ยกเว้นหัวจดหมาย (ที่อยู่ด้านบน)

**ตรวจกฎ** (MF / CNTRS ที่เป็น Excel) ถ้าเจอจะแจ้งเตือนบนหน้าเว็บ (กล่องใต้ตาราง) และในผลของคำสั่ง — `ผิด` = ผิดกฎ · `ตรวจ` = ให้คนตรวจซ้ำ:
SHED NO. (DG ลง LAEM CHABANG = 2826 · LAOS ลง BANGKOK = 0124 · USED ENGINE ลง BANGKOK = 0126 · SHED ใน MF ตรงกับ CNTRS) ·
ประเทศปลายทางตามข้อความ TRANSIT · CARGO MOVEMENT (TRANSHIPMENT = 3 · TRANSIT = 7) · TAX ID ตรงกับ NOTIFY ·
DG และขนาดตู้ใน MF ตรงกับ CNTRS · ตู้ DG ทุกตู้ใน CNTRS (รวม SOC) อยู่ SHED 2826

## เงื่อนไขการแยก DELIVERY (แสดงเฉพาะหัวข้อที่มีข้อมูล ไม่มีก็ข้ามไป)

| DELIVERY | มาจาก |
|---|---|
| THLCH / THLKR / THBMT / THSCT / THBKK | CNTRS: `Port Of Delivery` ของหน้านั้น ถ้าเป็นท่าเรือที่เรือลง ให้ดู `SHED NUMBER`: `0332` = THLKR · `0110` = THBMT · `0302` = THSCT · `0122` = THBKK · อื่น ๆ = THLCH (หรือ THBKK ถ้าเรือลงที่ BANGKOK) |
| MYANMAR / LAOS / MALAYSIA / CHINA / CAMBODIA | MF `3) CARGO MOVEMENT (7-TRANSIT)` และรหัสประเทศ `MM` / `LA` / `MY` / `CN` / `KH` |
| TRANSHIPMENT | MF CARGO MOVEMENT เป็น TRANSHIPMENT |
| EMPTY | ตู้เปล่า HEUNG A (STATUS `4=MTY`) ใน CNTRS ที่ไม่มีใน MF |

- **ตู้ที่อยู่ใน CNTRS แต่ไม่มีใน MF**: ของ HEUNG A ที่ไม่ใช่ตู้เปล่า นับรวมในแถว DELIVERY ของตู้นั้น · ตู้เปล่าเป็นแถว **EMPTY** · เจ้าของอื่นแสดงใต้ตารางในหัวข้อ **\*\*\*SOC\*\*\*** แยกตามเจ้าของ (คำแรกของ CONSIGNEE เช่น ORIENT, NS) และขนาด พร้อม `[SUM-EDI, MANIFEST AS PER ATTACHED]` แล้วรวมเป็น **GRAND TOTAL** (เอกสารไม่มีสี)
- SUMMARY.pdf: 1 หน้าต่อ 1 DELIVERY ที่มี B/L เรียงตามลำดับเดียวกับ SUM-EDI (THLCH = LAEM CHABANG · THLKR = LAEM CHABANG / LATKRABANG (GATE …) · THBMT = BANGKOK MODERN TERMINAL · THSCT = SIAM CONTAINER TERMINAL · THBKK = BANGKOK · Transit = \*CARGO INTRANSIT TO …) EMPTY และ SOC ไม่มีหน้า
- GATE ลาดกระบัง: ช่อง **GATE** บนหน้าเว็บ (หรือ `--gate`) เลือกได้ 2 แบบ `2` = LATKRABANG (ESCO GATE 2) (ค่าเริ่มต้น) · `8` = LATKRABANG (TIFFA GATE 8)
- หน้าเว็บ: ใส่ไฟล์ / ลายเซ็น / ชื่อ / GATE แล้วกดปุ่ม **Process** ด้านบน เว็บจะอ่านไฟล์และทำเอกสารใหม่ทั้งหมด (ถ้าแก้อะไรหลังกด จะขึ้นเตือนให้กด Process อีกครั้ง)
- จำนวน B/L นับรวม House B/L (…A, …B)
- ขนาดตู้จาก CNTRS ถ้าไม่มี ใช้รหัสประเภทตู้: `22G0` = 20' · `45G0` / `45R1` = 40' (ตู้เดียวกันนับครั้งเดียว)
- ยอดรวมทุกกลุ่มต้องเท่ากับ **GRAND TOTAL** ของ MF (โปรแกรมแจ้งว่าตรงหรือไม่ตรง)
- ท้าย SUM-EDI: ชื่อ / ตำแหน่ง / เบอร์โทร ของผู้จัดทำ (แก้ได้ เพราะมีพนักงานหลายคน)

## ใช้งานผ่านเว็บ (ไม่ต้องติดตั้งอะไร)

เปิดที่ **https://rattanao.github.io/Sum-edi/**

1. ลากไฟล์ MF และ CNTRS (Excel หรือ PDF หลายไฟล์ก็ได้) มาวางพร้อมกันในช่องเดียว เว็บแยกเองว่าไฟล์ไหนเป็น MF ไฟล์ไหนเป็น CNTRS แล้วแสดงตารางทันที (คลิกชื่อกลุ่มเพื่อดู B/L คลิกแถว SOC เพื่อดูเลขตู้)
2. ช่อง **ชื่อผู้จัดทำ / ตำแหน่ง / เบอร์โทร** แก้เป็นของตัวเอง เบราว์เซอร์จะจำไว้ในเครื่องนั้น
3. ช่อง **ลายเซ็น** เลือกรูปลายเซ็น (.png พื้นโปร่งใส หรือ .jpg) ครั้งเดียว กด **ลบลายเซ็น** เพื่อเอาออก
4. เลือก **LATKRABANG GATE** (2 = ESCO GATE 2 · 8 = TIFFA GATE 8) แล้วกด **Process**
5. ดูตัวอย่างทุกหน้าด้านล่าง (กด **ดูภาพเต็ม** หรือคลิกที่รูป เพื่อเปิดหน้าใหญ่ เลื่อนหน้าได้ด้วยปุ่ม ‹ › หรือลูกศรบนคีย์บอร์ด กด Esc เพื่อปิด) แล้วกด **ดาวน์โหลดทั้งหมด** หรือกดชื่อไฟล์ที่ต้องการ

ไฟล์ถูกประมวลผลในเบราว์เซอร์ของเครื่องที่ใช้อยู่เท่านั้น ไม่มีการอัปโหลดข้อมูล ชื่อ หรือลายเซ็นไปที่ใด ตัวเว็บคือไฟล์ [`index.html`](index.html)

## ใช้งานด้วยโปรแกรม Python (Windows)

สิ่งที่ต้องมี: Python 3.9 ขึ้นไป และฟอนต์ Times New Roman / Arial / Tahoma ของ Windows

```bash
pip install -r requirements.txt
```

1. วางไฟล์ MF (หรือ MANIFEST) และ CNTRS (หรือ INWARD) ไว้ในโฟลเดอร์ `input` เป็น Excel หรือ PDF ก็ได้ หลายไฟล์ก็ได้ (ใช้เฉพาะไฟล์ที่เป็นเที่ยวเรือเดียวกับ MF)
2. (ไม่บังคับ) วางรูปลายเซ็นชื่อ `assets/signature.png` (ไฟล์นี้ไม่ถูกส่งขึ้น GitHub)
3. ดับเบิลคลิก `run-sum-edi.bat` หรือสั่ง

   ```bash
   python -X utf8 scripts/sum_mf.py .
   ```

ผลลัพธ์อยู่ที่ `output/` : `SUM-EDI.pdf`, `SUMMARY.pdf`, `MF-*.pdf`, `CNTRS.pdf` (GATE ลาดกระบัง: `--gate 8`)

เปลี่ยนชื่อผู้จัดทำ: `python -X utf8 scripts/sum_mf.py . --name "MISS ..." --position "Document Inbound" --tel "02-059 ...."`
ระบุไฟล์เอง: `python -X utf8 scripts/sum_mf.py "D:\งาน\MANIFEST.xls" "D:\งาน\output" --cntrs "D:\งาน\CNTRS.xls"`

## ใช้กับ Claude Code

โฟลเดอร์นี้เป็น skill ชื่อ `sum-edi` (ไฟล์ [`SKILL.md`](SKILL.md), วางไว้ที่ `~/.claude/skills/sum-edi`) พิมพ์ `sum-edi` หรือ "ทำ SUM-EDI จาก MF" แล้วแนบไฟล์ MF และ CNTRS

---
Project by Ms. Rattana Mudlae · Import DOC. · HEUNG A LINE (THAILAND) CO.,LTD.
