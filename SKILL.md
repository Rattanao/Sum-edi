---
name: sum-edi
description: >
  "SUM-EDI": splits one full MF / MANIFEST (CARGO MANIFEST, Excel or PDF, for a whole vessel, the export
  that has the "1) SHED NO. / 2) country (XX) / 3) CARGO MOVEMENT" column) by DELIVERY, taken from the
  CNTRS container list (Excel or PDF, e.g. CNTRS.xls / INWARD.xls / CNTRS-A0.pdf, "CONTAINER LIST FOR FEEDER": Port Of Delivery + SHED
  NUMBER -> THLCH, THLKR, THBMT, THSCT, THBKK), the transit countries (MYANMAR, LAOS, MALAYSIA, CHINA,
  CAMBODIA), TRANSHIPMENT and EMPTY; only rows that have data are shown. Adds the SOC containers (in
  CNTRS but not in the manifest, other owners such as ORIENT / NS / BEN LINE) under ***SOC*** with a
  GRAND TOTAL; then writes output/SUM-EDI.pdf (A4 cover form, English) and output/SUMMARY.pdf (one
  signed landscape page per port / transit country / LKB), plus MF-<DELIVERY>.pdf (the Excel MF cut per
  DELIVERY, SUMMARY page as cover) and CNTRS.pdf (the Excel CNTRS reprinted per DELIVERY from the shed,
  e.g. 0332 -> THLKR /GATE 2). Files given = do all of it. Use whenever the user says "sum-edi",
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
- SUM-EDI form: AGENT NAME / TAX ID / VESSEL (as in the MF, e.g. `HEUNG-A BANGKOK 2612N`) / DISCHARGE (all four values start at the same x), table DELIVERY | TOTAL CNTRS (`n X 20'  n X 40'`) | TOTAL B/L | TOTAL GROSS WEIGHT | TOTAL PKGS, one blank row, `GRAND TOTAL n CNTRS` (`-` for zero B/L / weight). Footer `DATA SUBMITTED TO THE PORT BY ...` 1. EDI 2. E-MAIL and the staff block (name / position / TEL). Long text is drawn smaller to stay inside A4.
- SUMMARY.pdf (vessel written `NAME V.VOY`): one page per DELIVERY row that has B/L, in the SUM-EDI order. The port of discharge row (THLCH / THBKK) → `<PORT>`; transit countries → `*CARGO INTRANSIT TO <COUNTRY> P.D.R. VIA <PORT>`; any other row → `<PORT> / <full name>`: THLKR `LATKRABANG <gate>`, THBMT `BANGKOK MODERN TERMINAL`, THSCT `SIAM CONTAINER TERMINAL`, THBKK `BANGKOK`. Gate (`--gate`, web: GATE select, per job): `2` → `(ESCO GATE 2)` (default), `8` → `(TIFFA GATE 8)`; no other gates. EMPTY and SOC have no page.
- Staff block default MISS RATTANA / Document Inbound / 02-059 6253 (Direct); other staff: `--name`, `--position`, `--tel` (web: the three fields, kept per browser).
- Signature: `assets/signature.png` above the dotted line on every SUMMARY page. Local only (git-ignored, never pushed).
- **MF-<DELIVERY>.pdf** (`scripts/split_docs.py`, only when the MF is one Excel file): `MF-LCH.pdf`, `MF-LKR.pdf`, `MF-MYANMAR.pdf` … one per SUMMARY page, the B/L of that row laid out like the port system's CARGO MANIFEST print (Times 8, PORT TOTAL, then a GRAND TOTAL page), with that SUMMARY page in front as the cover. Red note from the shed on **every** page: THLKR `LKB (ESCO GATE 2)` / `LKB (TIFFA GATE 8)` (the gate) over the port; other terminals `LAEM CHABANG / BMT`, `/ SCT`. Reference SAWASDEE DENEB 2611S: LCH SUMMARY + 9 pages, LKR + 2, MYANMAR + 2, page breaks equal to the system prints.
- **CNTRS.pdf** (only when the CNTRS is Excel, which prints every shed under Port Of Delivery THLCH): the container list printed again like the system's per-delivery prints: one list per DELIVERY from the shed (0332 THLKR with red `/GATE n` on every page, 0110 THBMT, 0302 THSCT …) in the SUM-EDI order, then one list per SOC owner (ConTaxID left empty); a new page per SHED NUMBER (sheds in number order), items counted from 1 per shed, shed totals, Total Summary, As Agents, Page n / N, RecvControlNo / ARRIVAL DATE from the Excel. Reference DENEB: THLCH 24x20' 26x40' (transit 40' 1) 4 pages, THLKR 1x20' 3x40', LEGEND 1x20'.

- **Rule checks** (`scripts/checks.py`, Excel MF / CNTRS only; warnings, the documents are still made). From Check EDI: DG (CLASS / UN / HAZARD in the B/L) at LAEM CHABANG → SHED NO. 2826; LAOS at BANGKOK → 0124; USED ENGINE at BANGKOK → 0126; a TRANSIT sentence naming a country → that `2)` code (LAOS = LA, MYANMAR = MM …); TRANSHIPMENT → `(3-TRANSHIPMENT)`, TRANSIT (not Laos) → `(7-TRANSIT)`; a 3/7/L movement without any TRANSIT text → check; TAX ID name = NOTIFY name (the cut-off last word is allowed). MF against CNTRS: same SHED per container, DG in both or neither (per container, over all its B/L), same 20'/40'. CNTRS: every DG container (SOC too) at LAEM CHABANG under SHED 2826. DG outside 2826 is `ผิด` when the job uses SHED 2826, else `ตรวจ` (e.g. VEGA at terminal 2835). Reference: DENEB with LEGEND TIFU1013243 (DG tank) under 2820 → 1 `ผิด`; BKK 2612N → none.

## Steps

1. Find the files. With a job folder the script looks in `<job>/input`, then `<job>`. MF and CNTRS may each be Excel (.xls/.xlsx) or PDF, any file name, and may be split over several files (e.g. MF-LCH.pdf + MF-LAOS.pdf …, CNTRS-A0.pdf + CNTRS-LKR.pdf); they are told apart by content (`CONTAINER LIST FOR FEEDER` = CNTRS, `CARGO MANIFEST` = MF). Only files of the same voyage as the MF (e.g. 2611S) are used. `--cntrs FILE` gives one CNTRS file. PDF text is read line by line; the bigger page-header text is kept apart from container rows it is printed over.
2. Run:
   ```bash
   python -X utf8 "<skill dir>/scripts/sum_mf.py" "<job folder or MF file>" ["<output folder>"] [--cntrs FILE] [--gate 2|8] [--name ... --position ... --tel ...]
   ```
   Default output `<job>/output/`. Requires `pymupdf` and `xlrd`; Windows fonts Times New Roman, Arial and Tahoma.
3. Report to the user in Thai: vessel and port, the DELIVERY rows (20'/40', B/L, KGS, PKGS), GRAND TOTAL, the SOC lines and their GRAND TOTAL, whether the total = MF GRAND TOTAL, plus any warning the script prints (containers in the manifest but not in CNTRS, containers without a size, no signature, no CNTRS given).
   Also list the MF-*.pdf / CNTRS.pdf lines the script prints, and every `CHECKS` line (`[ผิด]` = wrong, `[ตรวจ]` = a person should look) with what to fix.
4. Send the user every output file: `SUM-EDI.pdf`, `SUMMARY.pdf`, `MF-*.pdf`, `CNTRS.pdf`. Files given = the user wants all of them; do not ask.

If the MF has no SHED NO. / country / CARGO MOVEMENT column (e.g. the per-port MF PDFs), the delivery comes from CNTRS and the transit country from the CNTRS REMARK `TRANSIT TO <COUNTRY>` or, failing that, from a B/L text that says TRANSIT and names the country; without CNTRS too, the script stops: ask for the CNTRS or the full MF export. Reference: SAWASDEE VEGA MF-*.pdf + CNTRS.xls gives the same result as the full MF..xls.

## Web tool

`index.html` does the same in the browser (SheetJS + pdf-lib + pdf.js, nothing uploaded):
https://rattanao.github.io/Sum-edi/ . MF and CNTRS go into one drop zone (told apart by content).
Signature and staff fields are kept per browser in localStorage. Nothing is built until the Process button (top of the page) is pressed; it re-reads every file and rebuilds all PDFs (SUM-EDI, SUMMARY, MF-*, CNTRS; the LATKRABANG GATE select
sets the red gate notes too) and shows the rule-check box under the table (red list, or ✓ when all pass), and any later change shows a reminder to press it again. On the web the Thai shed name in
CNTRS.pdf is drawn as a picture (the standard PDF fonts have no Thai). When the rules change, change
`scripts/sum_mf.py` + `scripts/split_docs.py` + `scripts/checks.py` and `index.html` together. Repo: https://github.com/Rattanao/Sum-edi (this skill
folder, `~/.claude/skills/sum-edi`, is its working copy).
