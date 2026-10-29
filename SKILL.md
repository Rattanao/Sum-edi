---
name: sum-edi
description: >
  "SUM-EDI": splits one full MF / MANIFEST (CARGO MANIFEST .xls/.xlsx for a whole vessel, the export
  that has the "1) SHED NO. / 2) country (XX) / 3) CARGO MOVEMENT" column) by DELIVERY, taken from the
  CNTRS container list (CNTRS.xls / INWARD.xls, "CONTAINER LIST FOR FEEDER": Port Of Delivery + SHED
  NUMBER -> THLCH, THLKR, THBMT, THSCT, THBKK), the transit countries (MYANMAR, LAOS, MALAYSIA, CHINA,
  CAMBODIA), TRANSHIPMENT and EMPTY; only rows that have data are shown. Adds the SOC containers (in
  CNTRS but not in the manifest, other owners such as ORIENT / NS / BEN LINE) under ***SOC*** with a
  GRAND TOTAL; then writes output/SUM-EDI.pdf (A4 cover form, English) and output/SUMMARY.pdf (one
  signed landscape page per port / transit country / LKB). Use whenever the user says "sum-edi",
  "SUM-EDI", "summary", "ทำ summary", "แยก MF", "แยกตาม Delivery", "ตู้ SOC", or points to an
  MF*/MANIFEST*.xls (with or without the CNTRS / INWARD .xls) and wants SUM-EDI.pdf / SUMMARY.pdf.
---

# SUM-EDI

Rattana's confirmed SUM-EDI / SUMMARY format (form sample: `C:\EDI\ตัวอย่างฟอร์ม.pdf`). Reference jobs:
- HEUNG-A BANGKOK 2612N (LAEM CHABANG): THLCH 28x20' 51x40' 53 B/L · THLKR 6x40' · THBMT 2x40' · THBKK
  2x20' 1x40' · EMPTY 150x20' · GRAND TOTAL 240 CNTRS / 59 B/L / 1,235,746.913 / 84,666 · SOC ORIENT 40' 5 → 245.
- SAWASDEE VEGA V.2613S (LAEM CHABANG): 63 CNTRS / 84 B/L / 896,471.042 KGS / 20,169 PKGS.
- KMTC BANGKOK V.2612S (BANGKOK): THBKK 17x20' 43x40' / 24 B/L, SOC BEN 20' 1 + NS 20' 21, GRAND TOTAL 82.

Run the bundled script; do not rebuild the layout by hand. All document text is English; only the
letterhead image at the top keeps its Thai address.

## Rules (already implemented in the script and in index.html)

DELIVERY row of each B/L, in this order (rows with no containers and no B/L are not shown):

| DELIVERY | Containers / B/L |
|---|---|
| THLCH, THLKR, THBMT, THSCT, THBKK | from CNTRS: `Port Of Delivery` of the page; when that equals the port of discharge (THLCH for LAEM CHABANG, THBKK for BANGKOK), the `SHED NUMBER` decides: `0332` THLKR, `0110` THBMT, `0302` THSCT, `0122` / `0145` THBKK, any other shed = the port of discharge. A B/L takes the delivery of its containers; a B/L whose containers are not in CNTRS falls back to its MF `SHED NO.` |
| MYANMAR, LAOS, MALAYSIA, CHINA, CAMBODIA | MF `3) CARGO MOVEMENT (7-TRANSIT)` and country `MM` / `LA` / `MY` / `CN` / `KH` |
| TRANSHIPMENT | MF CARGO MOVEMENT contains TRANSHIPMENT |
| EMPTY | HEUNG A containers in CNTRS, not in the MF, STATUS `4=MTY` (containers only) |

- **Containers in CNTRS but not in the manifest** (MF container numbers are read from column F or G):
  - HEUNG A, not empty → counted in their DELIVERY row.
  - HEUNG A empty (4=MTY) → EMPTY row.
  - Any other owner → under the table a `***SOC***` heading, one line per owner + size, owner = first word of the CONSIGNEE (`ORIENT OVERSEAS CONTAINER LINE` → ORIENT): `ORIENT  40'  5 CNTRS  [SUM-EDI, MANIFEST AS PER ATTACHED]`, then `GRAND TOTAL  <table total + SOC> CNTRS` (underlined). No colour in the PDF.
- B/L count includes house B/Ls (…A, …B). Container size from CNTRS (`(20')` / `(40')`), else the type code (`22G0` = 20', `45G0` = 40'). Containers counted once per row.
- Weight and packages = sum of the B/Ls in the row. The table's GRAND TOTAL must equal the MF **GRAND TOTAL**; the script prints OK / MISMATCH.
- SUM-EDI form: AGENT NAME / TAX ID / VESSEL (as in the MF, e.g. `HEUNG-A BANGKOK 2612N`) / PORT OF DISCHARGE, table DELIVERY | TOTAL CNTRS (`n X 20'  n X 40'`) | TOTAL B/L | TOTAL GROSS WEIGHT | TOTAL PKGS, one blank row, `GRAND TOTAL n CNTRS` (`-` for zero B/L / weight). Footer `DATA SUBMITTED TO THE PORT BY ...` 1. EDI 2. E-MAIL and the staff block (name / position / TEL). Long text is drawn smaller to stay inside A4.
- SUMMARY.pdf (vessel written `NAME V.VOY`): one page for the port (all Thai deliveries except THLKR together), one per transit country (`*CARGO INTRANSIT TO <COUNTRY> P.D.R. VIA <PORT>`), one for THLKR (`<PORT> / LKB (ESCO GATE 2)`). EMPTY, TRANSHIPMENT-only and SOC have no page of their own (TRANSHIPMENT B/Ls are in the port page).
- Staff block default MISS RATTANA / Import Documentation / 02-059 6253 (Direct); other staff: `--name`, `--position`, `--tel` (web: the three fields, kept per browser).
- Signature: `assets/signature.png` above the dotted line on every SUMMARY page. Local only (git-ignored, never pushed).

## Steps

1. Find the files. With a job folder the script looks in `<job>/input`, then `<job>`. The CNTRS list is found by content (`CONTAINER LIST FOR FEEDER`, any file name such as `CNTRS.xls` or `INWARD .xls`; `--cntrs FILE` to give it); the manifest is the other .xls, preferring a name starting with `MF`, then `MANIFEST`.
2. Run:
   ```bash
   python -X utf8 "<skill dir>/scripts/sum_mf.py" "<job folder or MF file>" ["<output folder>"] [--cntrs FILE] [--name ... --position ... --tel ...]
   ```
   Default output `<job>/output/`. Requires `pymupdf` and `xlrd`; Windows fonts Times New Roman, Arial and Tahoma.
3. Report to the user in Thai: vessel and port, the DELIVERY rows (20'/40', B/L, KGS, PKGS), GRAND TOTAL, the SOC lines and their GRAND TOTAL, whether the total = MF GRAND TOTAL, plus any warning the script prints (containers in the manifest but not in CNTRS, containers without a size, no signature, no CNTRS given).
4. Send the user `SUM-EDI.pdf` and `SUMMARY.pdf`.

If the MF has no SHED NO. / CARGO MOVEMENT column, the script stops: ask the user for the full MF export.

## Web tool

`index.html` does the same in the browser (SheetJS + pdf-lib + pdf.js, nothing uploaded):
https://rattanao.github.io/Sum-edi/ . MF and CNTRS go into one drop zone (told apart by content).
Signature and staff fields are kept per browser in localStorage. When the rules change, change
`scripts/sum_mf.py` and `index.html` together. Repo: https://github.com/Rattanao/Sum-edi (this skill
folder, `~/.claude/skills/sum-edi`, is its working copy).
