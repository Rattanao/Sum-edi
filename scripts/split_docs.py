"""Split documents, called by sum_mf.py after SUM-EDI.pdf / SUMMARY.pdf:
- MF-<DELIVERY>.pdf (MF-LCH.pdf, MF-LKR.pdf, MF-MYANMAR.pdf, ...): the full MANIFEST .xls cut per DELIVERY row, laid
  out like the port system's CARGO MANIFEST print, with that row's SUMMARY page in front as the cover. A red note
  from the shed goes on every page: THLKR "LKB (ESCO GATE 2)" over the port, other terminals "LAEM CHABANG / BMT".
- CNTRS.pdf: a CNTRS .xls (which prints every shed under Port Of Delivery THLCH) printed again as one CONTAINER LIST
  per DELIVERY (shed 0332 -> THLKR + red "/GATE n", 0110 -> THBMT, 0302 -> THSCT, ...) and one per SOC owner.
Only made from Excel inputs (MF / CNTRS PDFs are already the system's per-delivery prints).
"""
import os, re, datetime
import xlrd, pymupdf

import sum_mf as S

FONTS = {"a": r"C:\Windows\Fonts\arial.ttf", "ab": r"C:\Windows\Fonts\arialbd.ttf", "t": r"C:\Windows\Fonts\times.ttf",
         "an": r"C:\Windows\Fonts\ARIALN.TTF", "th": r"C:\Windows\Fonts\tahoma.ttf"}
RED = (1, 0.08, 0.1)
W, H = 841, 595
COLS = (35.0, 219.2, 389.4)       # B/L + parties | marks + containers | pkgs + description
RIGHT = 690.0                     # gross weight / measurement, right aligned
PITCH, GAP = 16.2, 18.4           # line pitch of a B/L, extra space before the next B/L
BOTTOM = 552.0                    # last baseline a B/L may use on a page
HEADER_ROWS = ("CARGO MANIFEST", "VESSEL & VOYAGE", "PORT OF LOADING", "BILL OF LANDING NO.", "1) SHIPPER",
               "3) NOTIFY PARTY")


def read_blocks(path):
    """MANIFEST .xls -> vessel, port of loading, port of discharge, {B/L: [rows of 4 texts]} (SHED column left out)"""
    s = xlrd.open_workbook(path, logfile=open(os.devnull, "w")).sheet_by_index(0)
    blocks, cur, vessel, pol, pod = {}, None, "", "", ""
    for r in range(s.nrows):
        v = [str(x).replace("\t", " ").rstrip() for x in s.row_values(r)] + [""] * 23
        line = " ".join(x for x in v if x.strip())
        if v[0] == "VESSEL & VOYAGE": vessel = vessel or v[2].strip()
        if v[0] == "PORT OF LOADING": pol, pod = pol or v[2].strip(), pod or v[8].strip()
        if "PORT TOTAL" in line or "GRAND TOTAL" in line: cur = None; continue
        if v[0].strip() in HEADER_ROWS or v[10].strip() == "CARGO MANIFEST" or v[0].startswith("PortOfDischarge"): continue
        if S.BLRE.match(v[0].strip()):
            cur = v[0].strip(); blocks[cur] = []
        if cur is None: continue
        row = [v[0].strip(), (v[5] or v[6]).strip(), v[10].strip(), v[16].strip()]
        if any(row): blocks[cur].append(row)
    return vessel, pol, pod, blocks


_FONT = {}


def tx(pg, x, y, s, f="a", size=8, right=False, center=False, color=(0, 0, 0)):
    if f not in _FONT: _FONT[f] = pymupdf.Font(fontfile=FONTS[f])
    w = _FONT[f].text_length(s, size)
    x = x - w if right else x - w / 2 if center else x
    pg.insert_text((x, y), s, fontname=f, fontfile=FONTS[f], fontsize=size, color=color)


class Pdf:
    def __init__(self, info):
        self.doc, self.info, self.pg, self.n = pymupdf.open(), info, None, 0

    def t(self, x, y, s, f="a", size=8, right=False, center=False, color=(0, 0, 0)):
        tx(self.pg, x, y, s, f, size, right, center, color)

    def dash(self, y):
        self.pg.draw_line((34.5, y), (816.8, y), width=0.5, dashes="[1.5 .5 .5 .5] 0")

    def new_page(self):
        """page with the vessel header; returns the first data baseline"""
        i = self.info
        self.pg = self.doc.new_page(width=W, height=H); self.n += 1
        d = 0.0
        if self.n == 1:
            self.pg.draw_rect((348.7, 18.0, 518.8, 46.3), width=0.5)
            self.t(433.7, 36.6, "CARGO MANIFEST", "ab", 13, center=True)
        else:
            d = -28.5
        # red note from the shed, on every page
        if i["note"]: self.t(337.8, 67.6 + d, i["note"], "an", 9, color=RED)            # LKB (ESCO GATE 2) above the port
        if i["pod_note"]:                                                                 # LAEM CHABANG / BMT
            self.t(339.0 + pymupdf.Font(fontfile=FONTS["a"]).text_length(i["pod"], 8) + 3, 81.9 + d, i["pod_note"], "t", 8.5, color=RED)
        for x, y, s, f in ((34.5, 66.2, "VESSEL & VOYAGE", "ab"), (123.2, 66.2, i["vessel"], "a"),
                           (440.8, 66.2, "NATIONALITY  OF SHIP", "ab"), (631.3, 67.8, "ARRIVAL DATE", "ab"),
                           (727.9, 67.8, "01/01/1900", "a"), (34.5, 81.9, "PORT OF LOADING", "ab"), (123.2, 81.9, i["pol"], "a"),
                           (239.2, 81.9, "PORT OF DISCHARGE", "ab"), (339.0, 81.9, i["pod"], "a"),
                           (442.4, 81.9, "NEXT PORT OF CALL", "ab"), (631.3, 81.9, "NAME OF MASTER", "ab")):
            self.t(x, y + d, s, f)
        top, bot = 89.7 + d, 132.2 + d
        for y in (top, bot): self.pg.draw_line((33.0, y), (815.7, y), width=0.5)
        for x in (33.0, 215.7, 382.6, 448.8, 595.2, 708.6, 815.7): self.pg.draw_line((x, top), (x, bot), width=0.5)
        self.pg.draw_line((708.6, top + 15.8), (815.7, top + 15.8), width=0.5)
        self.pg.draw_line((759.0, top + 15.8), (759.0, bot), width=0.5)
        self.t(39.2, 101.4 + d, "BILL OF LANDING NO.")
        for y, s in ((109.4, "1) SHIPPER"), (117.2, "2) CONSIGNEE"), (126.7, "3) NOTIFY PARTY")): self.t(39.2, y + d, s, size=7.3)
        self.t(299.2, 101.4 + d, "MARKS AND NUMBERS", center=True); self.t(299.2, 115.5 + d, "CONTAINER NO.", center=True)
        self.t(415.7, 115.5 + d, "NO. OF PKGS", center=True); self.t(522.0, 115.5 + d, "DESCRIPTION OF GOODS", center=True)
        self.t(599.8, 101.4 + d, "1) GROSS WEIGHT (KGM)"); self.t(599.8, 115.5 + d, "2) MEASUREMENT (MTQ)")
        self.t(762.2, 101.4 + d, "FREIGHT", center=True)
        self.t(733.8, 115.5 + d, "PREPAID", center=True); self.t(787.4, 115.5 + d, "COLLECT", center=True)
        self.t(38.2, 570.2, f"PortOfDischarge: {i['podc']} PlaceDelivery: {i['place']} RecvControlNo:  Print Date  : {i['date']}")
        self.t(747.2, 571.8, "Page", size=9); self.t(781.8, 571.8, str(self.n), size=9)
        return 145.2 + d

    def total(self, top, label, grey, pkg, gw, mtq):
        self.pg.draw_rect((33.0, top, 821.2, top + 50.1), color=None, fill=(grey,) * 3)
        self.pg.draw_line((33.0, top + 1.5), (815.2, top + 1.5), width=0.5, dashes="[1.5 .5 .5 .5] 0")
        self.t(165.0, top + 11.8, label, "ab", right=True)
        for dy, k, val in ((11.8, "TOTAL MEASUREMENT:", f"{mtq:,.3f} MTQ"), (26.1, "TOTAL GROSSWEIGHT:", f"{gw:,.3f} KGM"),
                           (42.5, "TOTAL PACK:", f"{pkg:,} PACKAGES")):
            self.t(276.0, top + dy, k, right=True); self.t(279.5, top + dy, val)


def write_group(path, info, nos, blocks, cover=None):
    """MF pages of one DELIVERY; cover = (SUMMARY doc, page no.) put in front as the cover sheet"""
    p = Pdf(info)
    y = p.new_page()
    pkg = gw = mtq = 0
    for no in nos:
        rows = blocks[no]
        first = rows[0]
        m = re.match(r"([\d,]+)\s", first[2]); pkg += int(m.group(1).replace(",", "")) if m else 0
        for r in rows:
            m = re.match(r"([\d,.]+)\s*(KGM|MTQ)", r[3])
            if m:
                if m.group(2) == "KGM": gw += S.num(m.group(1))
                else: mtq += S.num(m.group(1))
        if y + (len(rows) - 1) * PITCH > BOTTOM and y > 150: y = p.new_page()
        p.dash(y - 10.2)
        for r in rows:
            for x, s in zip(COLS, r[:3]):
                if s: p.t(x, y, s, "t")
            if r[3]: p.t(RIGHT, y, r[3], "t", right=True)
            y += PITCH
        y += GAP - PITCH          # next B/L baseline = last baseline + GAP
    last = y - GAP
    if last + 7.4 + 50.1 > BOTTOM + 8:
        p.new_page(); last = 116.7 - 7.4
    p.total(last + 7.4, "PORT TOTAL:", 0.878, pkg, gw, mtq)
    p.new_page()
    p.total(105.8, "GRAND TOTAL:", 0.753, pkg, gw, mtq)
    if cover: p.doc.insert_pdf(cover[0], from_page=cover[1], to_page=cover[1], start_at=0)
    p.doc.subset_fonts(); p.doc.save(path, garbage=3, deflate=True)
    return pkg, round(gw, 3), round(mtq, 3), p.n


# ---------------------------------------------------------------- CNTRS (CONTAINER LIST FOR FEEDER) per DELIVERY
HAL_TAX = "0993000456955"
ONES = ("ZERO ONE TWO THREE FOUR FIVE SIX SEVEN EIGHT NINE TEN ELEVEN TWELVE THIRTEEN FOURTEEN FIFTEEN SIXTEEN "
        "SEVENTEEN EIGHTEEN NINETEEN").split()
TENS = "- - TWENTY THIRTY FORTY FIFTY SIXTY SEVENTY EIGHTY NINETY".split()
CROW, CPITCH, CBOTTOM = 137.2, 19.8, 553.1     # first container baseline under SHED NUMBER, pitch, last baseline


def words(n):
    if n < 20: return ONES[n]
    if n < 100: return TENS[n // 10] + (f" {ONES[n % 10]}" if n % 10 else "")
    return f"{ONES[n // 100]} HUNDRED" + (f" {words(n % 100)}" if n % 100 else "")


def read_cntrs_rows(path):
    """CNTRS .xls -> header, [container rows with their SHED NUMBER]"""
    s = xlrd.open_workbook(path, logfile=open(os.devnull, "w")).sheet_by_index(0)
    head = dict(vessel="", voy="", arrival="", pod="", deliv="", recv="")
    rows, shed, name, cur = [], "", "", None
    for r in range(s.nrows):
        v = [str(x).strip() for x in s.row_values(r)] + [""] * 31
        if v[0] == "CONTAINER LIST FOR FEEDER" and not head["vessel"]:
            head.update(vessel=v[7], voy=v[12], arrival=v[17] or "__/__/____", pod=v[24], deliv=v[28])
        m = re.search(r"RecvControlNo:\s*(\S+)", " ".join(v))
        if m and not head["recv"]: head["recv"] = m.group(1)
        if v[0].startswith("SHED NUMBER"):
            m = re.match(r"\d+", v[2] or next((x for x in v[1:] if x), ""))
            shed, name, cur = (m.group(0).zfill(4) if m else ""), v[5], None
            continue
        if re.fullmatch(r"[A-Z]{4}\d{7}", v[0]):
            cur = dict(cntr=v[0], type=v[3], size=v[7], cons=[v[9]], gw=float(v[15] or 0), status=v[18], pol=v[20],
                       remark=" ".join(x for x in v[21:25] if x), shed=shed, shedname=name)
            rows.append(cur)
        elif cur and not v[0] and v[9] and not v[3] and not v[15]: cur["cons"].append(v[9])   # 2nd consignee line
        else: cur = None
    return head, rows


def cntrs_header(pg, head, deliv, note):
    tx(pg, 428.1, 33.4, "HEUNG A LINE (THAILAND) CO.,LTD.", "ab", 12, center=True)
    tx(pg, 428.1, 63.5, "TEL.02-0596200", size=9, center=True)
    for x, y, s, f in ((43.3, 80.4, "CONTAINER LIST FOR FEEDER", "ab"), (181.4, 80.4, head["vessel"], "a"),
                       (341.6, 81.0, "VOYAGE", "ab"), (382.2, 81.0, head["voy"], "a"),
                       (436.0, 81.0, "ARRIVAL DATE", "ab"), (509.0, 81.0, head["arrival"], "a"),
                       (573.1, 78.6, "Port Of Discharge", "ab"), (651.5, 78.6, head["pod"], "a"),
                       (693.5, 78.6, "Port Of Delivery", "ab"), (764.9, 78.6, deliv, "a")):
        tx(pg, x, y, s, f, 8.5)
    if note: tx(pg, 795.8, 80.0, note, "an", 10, color=RED)                    # THLKR /GATE 2
    for y in (92.4, 111.9): pg.draw_line((33.0, y), (816.0, y), width=0.5)
    for x, s in ((39.8, "CONTAINER NO."), (106.0, "ITEM"), (138.8, "TYPE"), (180.2, "SIZE"), (297.4, "CONSIGNEE"),
                 (437.0, "GROSS WEIGHT"), (520.5, "STATUS"), (571.5, "POL"), (615.5, "REMARK")):
        tx(pg, x, 104.4, s, "ab")


def cntrs_totals(pg, y, cs, summary=False, many=False):
    """Total Container 20' / 40' / Other, transhipment and transit (a '-' for none; 0 in the Total Summary)"""
    sz = lambda c, k: f"({k}')" in c["size"]
    n = {k: [c for c in cs if sz(c, k)] for k in ("20", "40")}
    other = len(cs) - len(n["20"]) - len(n["40"])
    zero = "0" if summary else "-"
    for dy, k in ((0, "20"), (17.2, "40")):
        cnt = len(n[k])
        tx(pg, 134.3, y + dy, f"Total Container {k}' :", right=True)
        tx(pg, 181.6, y + dy, str(cnt) if cnt else "-", "ab", right=True)
        if cnt: tx(pg, 187.4, y + dy, f"({words(cnt)} CONTAINER(S) ONLY)", "ab")
        ts = sum("TRANSHIP" in c["remark"].upper() for c in n[k])
        ti = sum("TRANSIT" in c["remark"].upper() and "TRANSHIP" not in c["remark"].upper() for c in n[k])
        tx(pg, 421.0, y + dy, f"Total Transhipment Con. {k}' :"); tx(pg, 574.9, y + dy, str(ts) if ts else zero, "ab", right=True)
        tx(pg, 600.8, y + dy, f"Total Transit Con. {k}' :"); tx(pg, 730.8, y + dy, str(ti) if ti else zero, "ab", right=True)
    tx(pg, 134.2, y + 34.3, "Total Container Other :", right=True)
    tx(pg, 181.6, y + 34.3, str(other) if other else "-", "ab", right=True)
    if other or (summary and many): tx(pg, 187.4, y + 34.3, f"({words(other) + ' ' if other else ''}CONTAINER(S) ONLY)", "ab")


def cntrs_section(doc, head, deliv, sheds, tax, recv, date, note):
    """one CONTAINER LIST: a new page per SHED NUMBER, shed totals, Total Summary, As Agents; Page n / N"""
    pages = []

    def page():
        pg = doc.new_page(width=W, height=H); pages.append(pg.number)
        cntrs_header(pg, head, deliv, note)
        tx(pg, 817.6, 571.5, f"ConTaxID: {tax} PortOfDischarge: {head['pod']} PortOfLoading: __/__/____ PlaceDelivery: "
                             f"{deliv} ShedNumber: 1 RecvControlNo: {recv} Print Date  : {date}", right=True)
        return pg

    for shed, name, cs in sheds:
        pg = page()
        tx(pg, 38.0, 123.2, "SHED NUMBER:", "ab"); tx(pg, 125.2, 123.2, shed)
        if name: tx(pg, 148.4, 123.2, name, "th", 7.5)
        y = CROW
        for i, c in enumerate(cs, 1):
            if y > CBOTTOM: pg = page(); y = 130.5
            tx(pg, 40.3, y, c["cntr"]); tx(pg, 113.6, y, str(i), center=True); tx(pg, 127.2, y, c["type"])
            tx(pg, 171.9, y, c["size"])
            for k, line in enumerate(c["cons"][:2]): tx(pg, 215.7, y + 9.2 * k, line)
            tx(pg, 489.9, y, f"{c['gw']:,.3f}", right=True); tx(pg, 493.9, y, "KGM")
            tx(pg, 518.7, y, c["status"]); tx(pg, 563.0, y, c["pol"]); tx(pg, 598.4, y, c["remark"])
            y += CPITCH
        lab = y - CPITCH + 30.1
        if lab + 34.3 > 556: pg = page(); lab = 123.1
        cntrs_totals(pg, lab, cs)
    if lab + 157.4 > 556: pg = page(); lab = 80.0
    pg.draw_line((34.8, lab + 46.3), (817.8, lab + 46.3), width=0.5)
    tx(pg, 50.5, lab + 60.1, "Total Summary", "ab")
    cntrs_totals(pg, lab + 74.1, [c for _, _, cs in sheds for c in cs], summary=True, many=len(sheds) > 1)
    tx(pg, 452.1, lab + 139.8, "HEUNG A LINE (THAILAND) CO.,LTD.", "ab", 9); tx(pg, 578.2, lab + 157.4, "As Agents")
    for i, pno in enumerate(pages, 1):
        pg = doc[pno]
        tx(pg, 747.2, 35.1, "Page", size=9); tx(pg, 785.7, 35.1, str(i), size=9, right=True)
        tx(pg, 789.2, 35.7, "/", size=10); tx(pg, 793.2, 35.1, str(len(pages)), size=9)
    return len(pages)


def write_cntrs(path, cn_file, gate, date):
    """CNTRS .xls -> CNTRS.pdf: HEUNG A containers per DELIVERY (from the shed) in the sum-edi order, then each SOC
    owner; inside a list the sheds in number order, items counted again from 1 per shed"""
    head, rows = read_cntrs_rows(cn_file)
    rank = {g: i for i, g in enumerate(S.ORDER)}
    sect = {}
    for c in rows:
        d = S.SHED_DELIVERY.get(c["shed"], head["deliv"]) if head["deliv"] == head["pod"] else head["deliv"]
        cons = " ".join(c["cons"])
        own = "" if S.HALRE.search(cons) else S.soc_owner(cons)
        sect.setdefault((own != "", rank.get(d, 99), d, own), {}).setdefault((c["shed"], c["shedname"]), []).append(c)
    doc, out = pymupdf.open(), []
    for (soc, _, d, own), sh in sorted(sect.items()):
        sheds = [(s, nm, cs) for (s, nm), cs in sorted(sh.items(), key=lambda x: int(x[0][0] or 0))]
        n = cntrs_section(doc, head, d, sheds, "" if soc else HAL_TAX, head["recv"], date,
                          f"/GATE {gate}" if d == "THLKR" and not soc else "")
        cnt = sum(len(cs) for _, _, cs in sheds)
        out.append(f"{own or 'HEUNG A'} {d}: {cnt} CNTRS, {n} page(s)")
    doc.subset_fonts(); doc.save(path, garbage=3, deflate=True)
    return out


def write_all(out, mf, cn, main_code, summ, summary_pdf, gate):
    """MF-<DELIVERY>.pdf (from one Excel MF; covers = the pages of SUMMARY.pdf, one per row with B/L in summ order)
    and CNTRS.pdf (from an Excel CNTRS) -> lines to print"""
    now = datetime.datetime.now()
    date = now.strftime(f"%d/%m/{now.year + 543} %H:%M:%S")
    lines = []
    xls = [f for f in mf if re.search(r"\.xlsx?$", f, re.I)]
    if len(mf) == 1 and xls:
        vessel, pol, pod, blocks = read_blocks(xls[0])
        rows = [x for x in summ if x["bls"] and not x.get("empty")]
        sdoc = pymupdf.open(summary_pdf)
        for i, x in enumerate(rows):
            g = x["g"]
            name = re.sub(r"^TH(?=[A-Z]{3}$)", "", g)
            terminal = re.fullmatch(r"TH[A-Z]{3}", g) and g not in (main_code, "THLKR")
            info = dict(vessel=vessel, pol=pol, pod=pod, podc=main_code, date=date,
                        place=g if re.fullmatch(r"TH[A-Z]{3}", g) else main_code,
                        note=f"LKB {S.gate_text(gate)}" if g == "THLKR" else "", pod_note=f"/ {name}" if terminal else "")
            nos = [n for n in x["nos"] if n in blocks]
            pkg, gw, mtq, n = write_group(os.path.join(out, f"MF-{name}.pdf"), info, nos, blocks, (sdoc, i))
            lines.append(f"MF-{name}.pdf  SUMMARY + {n} pages  {len(nos)} B/L  {pkg:,} PKGS  {gw:,.3f} KGM  {mtq:,.3f} MTQ")
        sdoc.close()
    cx = [f for f in cn if re.search(r"\.xlsx?$", f, re.I)]
    if cx:
        for line in write_cntrs(os.path.join(out, "CNTRS.pdf"), cx[0], gate, date): lines.append(f"CNTRS.pdf  {line}")
    return lines
