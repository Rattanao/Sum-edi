---
name: summary
description: >
  "SUMMARY" / SUM-EDI: splits one full MF / MANIFEST (CARGO MANIFEST .xls/.xlsx for a whole vessel, the
  export that has the "1) SHED NO. / 2) country (XX) / 3) CARGO MOVEMENT" column) into the main port
  (LAEM CHABANG = LCH or BANGKOK, from PORT OF DISCHARGE), LKB (GATE2) (SHED NO. 0332) and the transit
  countries MYANMAR / LAOS / MALAYSIA (CARGO MOVEMENT 7-TRANSIT, country MM / LA / MY); adds the SOC
  containers (in CNTRS.xls but not in the manifest, e.g. BEN LINE / NS) as own lines and a GRAND
  TOTAL; then writes output/SUM-EDI.pdf (A4 cover table, English) and output/SUMMARY.pdf (one signed
  landscape page per group). Use whenever the user says "summary", "sum-edi", "SUM-EDI", "ทำ summary",
  "แยก MF", "แยกย่อยเป็นไฟล์ SUM-EDI / SUMMARY", "ตู้ SOC", or points to an MF*/MANIFEST*.xls (with or
  without CNTRS.xls) and wants SUM-EDI.pdf / SUMMARY.pdf.
---

# SUMMARY

Rattana's confirmed SUM-EDI / SUMMARY format. Reference jobs: SAWASDEE VEGA V.2613S (LAEM CHABANG,
63 CNTRS / 84 B/L / 896,471.042 KGS / 20,169 PKGS) and KMTC BANGKOK V.2612S (BANGKOK, 17x20' 43x40' /
24 B/L / 1,223,789.730 / 49,442, SOC BEN LINE 20' 1 + NS 20' 21, GRAND TOTAL 82 CNTRS).
Run the bundled script; do not rebuild the layout by hand. All document text is English; only the
letterhead image at the top keeps its Thai address.

## Rules (already implemented in the script and in index.html)

| Group | B/L in the group | Label in SUM-EDI | SUMMARY page |
|---|---|---|---|
| Main port | every B/L not in another group; port = `PORT OF DISCHARGE` of the MF | `LCH` (LAEM CHABANG) / `BANGKOK` | `<PORT>` (agent HEUNG A LINE CO., LTD. C/O HEUNG A LINE (THAILAND) CO.,LTD.) |
| LKB | `1) SHED NO. 0332` | `LKB (GATE2)` | `<PORT> /` LKB (ESCO GATE 2) |
| Transit | `3) CARGO MOVEMENT (7-TRANSIT)` and country in `2) ... (XX)` is not TH | `MYANMAR` (MM), `LAOS` (LA), `MALAYSIA` (MY) | \*CARGO INTRANSIT TO <COUNTRY> P.D.R. VIA <PORT> (agent HEUNG A LINE CO.,LTD C/O HEUNG-A SHIPPING (THAILAND) CO.,LTD.) |

- SUMMARY title: `SUMMARY OF INWARD CARGO MANIFEST DISCHARGING AT <PORT>`. Page order: main port, transit countries, LKB.
- B/L count includes house B/Ls (…A, …B).
- Container size: from CNTRS (`(20')` / `(40')`) when given, else the B/L type code (`22G0` = 20', `45G0` / `45R1` = 40'). Containers are counted once per group.
- **Containers in CNTRS but not in the manifest** (MF container numbers are read from column F or G):
  - HEUNG A (consignee HEUNG A …), not empty → counted in the main port row (LCH / BANGKOK) as normal.
  - HEUNG A empty (CNTRS STATUS `4=MTY`) → an `EMPTY  n X 20'` row inside the table (containers only, no B/L / weight / packages), included in TOTAL.
  - Any other owner → under TOTAL a `***SOC***` heading, then one line per owner + size, owner = first word of the CONSIGNEE (`ORIENT OVERSEAS CONTAINER LINE` → ORIENT, `NS INTERNATIONAL …` → NS): `ORIENT  40' :  5 CNTRS  [SUM-EDI, MANIFEST AS PER ATTACHED]` (no colour in the PDF).
  - Then `GRAND TOTAL  <TOTAL + SOC> CNTRS` (bold, underlined). EMPTY and SOC are not on SUMMARY. Text that would run past the A4 margins (long owner, vessel or staff names) is drawn smaller so it fits.
- Weight and packages = sum of the B/Ls in the group. TOTAL must equal the MF **GRAND TOTAL**; the script prints OK / MISMATCH.
- SUM-EDI table: at least 4 rows and always one ruled blank row before TOTAL. Headers: NO. OF CONTAINERS | NO. OF B/L (D/O) | TOTAL GROSS WEIGHT (KGS) | TOTAL NO. OF PACKAGES. Footer: `DATA SUBMITTED TO THE PORT BY ...` 1. EDI 2. E-MAIL, then the staff block (name / position / TEL).
- Staff block: default MISS RATTANA / Import Documentation / 02-059 6253 (Direct). Other staff: `--name`, `--position`, `--tel` (web: the three fields, kept per browser).
- Signature: `assets/signature.png` above the dotted line on every SUMMARY page. Local only (git-ignored, never pushed).
- Vessel: `VESSEL & VOYAGE` written `NAME V.VOY` (KMTC BANGKOK 2612S → KMTC BANGKOK V.2612S).

## Steps

1. Find the files. With a job folder the script looks in `<job>/input`, then `<job>`: the manifest is the .xls whose name starts with `MF`, else `MANIFEST`, else any other .xls not starting with `CNTRS`; CNTRS is the .xls starting with `CNTRS` (optional, `--cntrs FILE` to give it).
2. Run:
   ```bash
   python -X utf8 "<skill dir>/scripts/sum_mf.py" "<job folder or MF file>" ["<output folder>"] [--cntrs FILE] [--name ... --position ... --tel ...]
   ```
   Default output `<job>/output/`. Requires `pymupdf` and `xlrd`; Windows fonts Times New Roman and Tahoma.
3. Report to the user in Thai: vessel and port, the table per group (20'/40', B/L, KGS, PKGS), TOTAL, the SOC lines and GRAND TOTAL, whether TOTAL = MF GRAND TOTAL, plus any warning the script prints (containers in the manifest but not in CNTRS, containers without a size, no signature, no CNTRS given = no SOC check).
4. Send the user `SUM-EDI.pdf` and `SUMMARY.pdf`.

If the MF has no SHED NO. / CARGO MOVEMENT column, the script stops: ask the user for the full MF export.

## Web tool

`index.html` does the same in the browser (SheetJS + pdf-lib + pdf.js, nothing uploaded):
https://rattanao.github.io/Sum-edi/ . MF and CNTRS go into one drop zone (told apart by content).
Signature and staff fields are kept per browser in localStorage. When the rules change, change
`scripts/sum_mf.py` and `index.html` together. Repo: https://github.com/Rattanao/Sum-edi (this skill
folder is its working copy).
