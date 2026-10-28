---
name: summary
description: >
  "SUMMARY" / SUM-EDI: splits one full MF (CARGO MANIFEST .xls/.xlsx for a whole vessel, the export
  that has the "1) SHED NO. / 2) country (XX) / 3) CARGO MOVEMENT" column) into LCH, LKB (GATE2)
  (SHED NO. 0332) and the transit countries MYANMAR / LAOS / MALAYSIA (CARGO MOVEMENT 7-TRANSIT,
  country MM / LA / MY), then writes output/SUM-EDI.pdf (A4 cover table: containers 20'/40', B/L
  count, gross weight, packages per group + TOTAL) and output/SUMMARY.pdf (one landscape "SUMMARY OF
  INWARD CARGO MANIFEST DISCHARGING AT LAEM CHABANG" page per group, signed). Use whenever the user
  says "summary", "sum-edi", "ทำ summary", "แยก MF", "แยกย่อยเป็นไฟล์ SUM-EDI / SUMMARY", or points to
  an MF*.xls and wants SUM-EDI.pdf / SUMMARY.pdf.
---

# SUMMARY

Rattana's confirmed SUM-EDI / SUMMARY format (vessel SAWASDEE VEGA V.2613S was the reference job).
Run the bundled script; do not rebuild the layout by hand.

## Rules (already implemented in the script and in index.html)

| Group | B/L in the group | Label in SUM-EDI | SUMMARY page |
|---|---|---|---|
| LCH | every B/L not in another group | `LCH` | LAEM CHABANG (agent HEUNG A LINE CO., LTD. C/O HEUNG A LINE (THAILAND) CO.,LTD.) |
| LKB | `1) SHED NO. 0332` | `LKB (GATE2)` | LAEM CHABANG / LKB (ESCO GATE 2) |
| Transit | `3) CARGO MOVEMENT (7-TRANSIT)` and country in `2) ... (XX)` is not TH | `MYANMAR` (MM), `LAOS` (LA), `MALAYSIA` (MY) | \*CARGO INTRANSIT TO <COUNTRY> P.D.R. VIA LAEM CHABANG (agent HEUNG A LINE CO.,LTD C/O HEUNG-A SHIPPING (THAILAND) CO.,LTD.) |

- B/L count includes house B/Ls (…A, …B).
- Containers are counted once per group; size from the B/L's type code: `22G0` = 20', `45G0` / `45R1` (and any 4x code) = 40'.
- Weight and packages = sum of the B/Ls in the group. The total of all groups must equal the MF **GRAND TOTAL**; the script prints OK / MISMATCH.
- SUMMARY page order: LCH, the transit countries (MYANMAR, LAOS, MALAYSIA), then LKB.
- Signature: `assets/signature.png` is put above the dotted line on every SUMMARY page. It is local only (git-ignored, never pushed). Without it the pages are unsigned.
- Vessel: `VESSEL & VOYAGE` of the MF, written `NAME V.VOY` (SAWASDEE VEGA 2613S → SAWASDEE VEGA V.2613S).

## Steps

1. Find the MF file. The script looks in `<job>/input`, then `<job>`, preferring a file whose name starts with `MF`; a file path also works.
2. Run:
   ```bash
   python -X utf8 "<skill dir>/scripts/sum_mf.py" "<job folder or MF file>" ["<output folder>"]
   ```
   The default output is `<job>/output/` (for a file path: `output/` next to the file). Requires `pymupdf` and `xlrd`; Windows fonts Times New Roman and Tahoma.
3. Report to the user in Thai: the vessel, the table per group (20'/40', B/L, KGS, PKGS) and TOTAL, whether TOTAL = MF GRAND TOTAL, plus any warning the script prints (SHED other than 2835 counted as LCH, containers without a size, no signature).
4. Send the user `SUM-EDI.pdf` and `SUMMARY.pdf`.

If the MF has no SHED NO. / CARGO MOVEMENT column (for example the plain MANIFEST.xls), the script stops with a message: ask the user for the full MF export.

## Web tool

`index.html` does the same in the browser (SheetJS + pdf-lib + pdf.js, nothing uploaded):
https://rattanao.github.io/Sum-edi/ . The signature is chosen once per browser and kept in localStorage.
When the rules change, change `scripts/sum_mf.py` and `index.html` together. The repo is
https://github.com/Rattanao/Sum-edi (this skill folder is its working copy).
