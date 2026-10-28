"""SUMMARY: split one full MF .xls (CARGO MANIFEST with the SHED NO. / country / CARGO MOVEMENT
column) into LCH, LKB (GATE2) and the transit countries, and write SUMMARY.pdf + SUM-EDI.pdf.

usage: python sum_mf.py [job folder | MF file] [output folder]
  job folder (default: current folder): MF is looked up in <job>/input, then <job>
  output folder default: <job>/output
"""
import sys, os, re, glob
import xlrd, pymupdf

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "..", "assets")
LETTERHEAD = os.path.join(ASSETS, "letterhead.jpg")
SIGNATURE = os.path.join(ASSETS, "signature.png")  # local only, not in git

BLRE = re.compile(r"^HASL[A-Z0-9]{10,14}$")
CNRE = re.compile(r"\b([A-Z]{4}\d{7})\b")
num = lambda t: float(t.replace(",", ""))
COUNTRY = {"MM": "MYANMAR", "LA": "LAOS", "MY": "MALAYSIA", "KH": "CAMBODIA", "VN": "VIETNAM", "CN": "CHINA"}
SHED_GROUP = {"0332": "LKR"}  # any other shed = LCH
LABEL = {"LCH": "LCH", "LKR": "LKB (GATE2)"}
ORDER = ["LCH", "LKR", "MYANMAR", "LAOS", "MALAYSIA"]


def find_mf(arg):
    if os.path.isfile(arg): return arg, os.path.join(os.path.dirname(os.path.abspath(arg)), "output")
    for d in (os.path.join(arg, "input"), arg):
        files = [f for f in glob.glob(os.path.join(d, "*.xls*")) if not os.path.basename(f).startswith("~$")]
        files.sort(key=lambda f: (not os.path.basename(f).upper().startswith("MF"), f))
        if files: return files[0], os.path.join(arg, "output")
    sys.exit(f"!! ไม่พบไฟล์ MF (.xls/.xlsx) ใน {arg}\\input หรือ {arg}")


def read_mf(path):
    s = xlrd.open_workbook(path, logfile=open(os.devnull, "w")).sheet_by_index(0)
    bls, cur, vessel, gt = {}, None, "", None
    for r in range(s.nrows):
        v = [str(x).strip() for x in s.row_values(r)] + [""] * 23
        line = " ".join(x for x in v if x)
        if v[0] == "VESSEL & VOYAGE": vessel = v[2]
        if "GRAND TOTAL" in line: gt = r
        if BLRE.match(v[0]):
            cur = v[0]
            bls[cur] = dict(pkg=0, gw=0.0, shed="", country="", transit=False, cntrs=[], types=[])
        if cur is None: continue
        if v[0].startswith("PortOfDischarge") or "PORT TOTAL" in line or "GRAND TOTAL" in line:
            cur = None; continue
        b = bls[cur]
        m = re.match(r"^([\d,]+)\s+[A-Z]{2}\b", v[10])
        if m and not b["pkg"]: b["pkg"] = int(num(m.group(1)))
        m = re.match(r"^([\d,.]+) KGM", v[16])
        if m and not b["gw"]: b["gw"] = num(m.group(1))
        m = re.search(r"SHED NO\.\s*(\d+)", v[15])
        if m: b["shed"] = m.group(1)
        m = re.match(r"^2\).*\(([A-Z]{2})\)\s*$", v[15])
        if m: b["country"] = m.group(1)
        if "CARGO MOVEMENT" in v[15] and "TRANSIT" in v[15]: b["transit"] = True
        b["cntrs"] += CNRE.findall(v[5])
        m = re.match(r"^(\d{2}[A-Z]\d)X\d+", v[10])
        if m: b["types"].append(m.group(1))
    total = None
    if gt is not None:
        blk = " ".join(" ".join(str(x) for x in s.row_values(r)) for r in range(gt, min(gt + 4, s.nrows)))
        m = re.search(r"GROSSWEIGHT:\s*([\d,.]+).*?TOTAL PACK:\s*([\d,]+)", blk)
        if m: total = (num(m.group(1)), int(num(m.group(2))))
    return vessel, bls, total


def group_of(b):
    if b["transit"] and b["country"] not in ("", "TH"): return COUNTRY.get(b["country"], b["country"])
    return SHED_GROUP.get(b["shed"], "LCH")


def summarize(bls):
    size = {}
    for b in bls.values():
        for c in b["cntrs"]:
            if b["types"]: size.setdefault(c, "20'" if b["types"][0][0] == "2" else "40'")
    groups = {}
    for no, b in bls.items(): groups.setdefault(group_of(b), []).append(no)
    order = ORDER + [g for g in groups if g not in ORDER]
    summ = []
    for g in [g for g in order if g in groups]:
        nos = groups[g]
        cn = {c for n in nos for c in bls[n]["cntrs"]}
        summ.append(dict(g=g, nos=nos, bls=len(nos), cntrs=len(cn),
                         n20=sum(size.get(c) == "20'" for c in cn), n40=sum(size.get(c) == "40'" for c in cn),
                         gw=round(sum(bls[n]["gw"] for n in nos), 3), pkg=sum(bls[n]["pkg"] for n in nos)))
    nosize = sorted(c for c in {c for b in bls.values() for c in b["cntrs"]} if c not in size)
    return summ, nosize


# ---------------------------------------------------------------- drawing
TIMES, TIMESB = r"C:\Windows\Fonts\times.ttf", r"C:\Windows\Fonts\timesbd.ttf"
FONTFILES = {"tr": TIMES, "tb": TIMESB, "th": r"C:\Windows\Fonts\tahoma.ttf", "thb": r"C:\Windows\Fonts\tahomabd.ttf"}
AGENT1 = "HEUNG A LINE CO., LTD. C/O HEUNG A LINE (THAILAND) CO.,LTD."
AGENT2 = "HEUNG A LINE CO.,LTD C/O HEUNG-A SHIPPING (THAILAND) CO.,LTD."


def text(pg, x, y, s, size=8, bold=False, right=False, underline=False, font=None):
    fn = font or ("tb" if bold else "tr"); ff = FONTFILES[fn]
    w = pymupdf.Font(fontfile=ff).text_length(s, size)
    if right: x -= w
    pg.insert_text((x, y), s, fontsize=size, fontname=fn, fontfile=ff)
    if underline: pg.draw_line((x, y + 1.5), (x + w, y + 1.5), width=0.5)
    return w


def summary_page(doc, vessel_name, agent, left, place, gw, pkg, double=False):
    pg = doc.new_page(width=842, height=595)
    text(pg, 421 - 190, 60, "SUMMARY OF INWARD CARGO MANIFEST DISCHARGING AT LAEM CHABANG", 8, bold=True)
    text(pg, 133, 99, "FOR MV.", 8, underline=True); text(pg, 177, 99, vessel_name, 8, underline=True)
    pg.draw_line((75, 118), (690, 118), width=0.6, dashes="[1 1.5] 0")
    text(pg, 75, 178 if not double else 198, agent, 8, underline=True)
    y = 222 if not double else 238
    text(pg, 75, y, left, 8)
    if place: text(pg, 222, y, place, 8)
    for yy in (y + 4, y + 40):
        text(pg, 558, yy, f"{gw:,.3f}", 8.5, right=True); text(pg, 566, yy, "KGM", 8)
        text(pg, 666, yy, f"{pkg:,}", 8.5, right=True); text(pg, 674, yy, "PACKAGES", 8)
    text(pg, 433, y + 38, "TOTAL", 7.5)
    pg.draw_line((470, y + 26), (738, y + 26), width=0.7)
    pg.draw_line((470, y + 45), (738, y + 45), width=0.7)
    if double: pg.draw_line((470, y + 47), (738, y + 47), width=0.7)
    if os.path.exists(SIGNATURE): pg.insert_image(pymupdf.Rect(88, 358, 88 + 27, 358 + 44.5), filename=SIGNATURE)
    pg.draw_line((75, 404), (690, 404), width=0.6, dashes="[1 1.5] 0")
    text(pg, 75, 424, "FOR HEUNG A LINE (THAILAND) CO.,LTD.", 7.5)


def write_summary(path, vessel_name, summ):
    S = {x["g"]: x for x in summ}
    doc = pymupdf.open()
    if "LCH" in S: summary_page(doc, vessel_name, AGENT1, "LAEM CHABANG", "", S["LCH"]["gw"], S["LCH"]["pkg"], double=True)
    for g in [g for g in S if g not in ("LCH", "LKR")]:
        summary_page(doc, vessel_name, AGENT2, "*CARGO INTRANSIT TO",
                     f"{g}{'  ' if g == 'LAOS' else ''}   P.D.R.  VIA  LAEM CHABANG", S[g]["gw"], S[g]["pkg"])
    if "LKR" in S: summary_page(doc, vessel_name, AGENT1, "LAEM CHABANG /", "LKB (ESCO GATE 2)", S["LKR"]["gw"], S["LKR"]["pkg"], double=True)
    doc.subset_fonts(); doc.save(path, garbage=4, deflate=True)


def write_sum_edi(path, vessel_name, summ, T):
    doc = pymupdf.open(); pg = doc.new_page(width=595, height=842)
    pg.insert_image(pymupdf.Rect(45, 52, 45 + 505, 52 + 102), filename=LETTERHEAD)
    for i, (k, v) in enumerate((("AGENT  NAME  :", AGENT1), ("TAX ID:", "0993000456955"), ("VESSEL  :", vessel_name))):
        y = 194 + i * 21
        text(pg, 53, y, k, 9.5, font="thb"); text(pg, 147, y, v, 9.5, font="thb")
        pg.draw_line((140, y + 5), (510, y + 5), width=0.5, color=(0.5, 0.5, 0.5))
    X = [140, 263, 330, 420, 520]
    top, hh, rh = 270, 50, 20
    nrows = len(summ) + 1
    bottom = top + hh + rh * nrows + 22
    pg.draw_rect(pymupdf.Rect(X[0], top, X[-1], top + hh), width=0.7)
    for x in X: pg.draw_line((x, top), (x, bottom), width=0.7)

    def ctext(x0, x1, y, s, size=9, font="th"):
        w = pymupdf.Font(fontfile=FONTFILES[font]).text_length(s, size); text(pg, (x0 + x1 - w) / 2, y, s, size, font=font)
    ctext(X[0], X[1], top + 30, "จำนวนตู้"); ctext(X[1], X[2], top + 30, "จำนวน B/L")
    ctext(X[2], X[3], top + 20, "น้ำหนักสินค้า"); ctext(X[2], X[3], top + 40, "รวม(KGS)")
    ctext(X[3], X[4], top + 20, "จำนวนหีบห่อ"); ctext(X[3], X[4], top + 40, "สินค้ารวม")
    y = top + hh + 15
    for x in summ:
        text(pg, X[0] - 2, y, LABEL.get(x["g"], x["g"]), 8.5, right=True, font="thb")
        if x["n20"]: text(pg, 185, y, f"{x['n20']}  X 20'", 8.5, right=True, font="th")
        if x["n40"]: text(pg, 258, y, f"{x['n40']}  X 40'", 8.5, right=True, font="th")
        text(pg, 300, y, f"{x['bls']}", 8.5, right=True, font="th")
        text(pg, 412, y, f"{x['gw']:,.3f}", 8.5, right=True, font="th")
        text(pg, 514, y, f"{x['pkg']:,}", 8.5, right=True, font="th")
        y += rh
    yt = top + hh + rh * nrows + 15
    pg.draw_line((X[0], yt - 15), (X[-1], yt - 15), width=0.7); pg.draw_line((X[0], yt + 7), (X[-1], yt + 7), width=0.7)
    text(pg, X[0] - 2, yt, "TOTAL", 8.5, right=True, font="thb")
    text(pg, 205, yt, f"{T['cntrs']}", 8.5, right=True, font="th"); text(pg, 215, yt, "CNTRS", 8.5, font="th")
    text(pg, 300, yt, f"{T['bls']}", 8.5, right=True, font="th")
    text(pg, 412, yt, f"{T['gw']:,.3f}", 8.5, right=True, font="th")
    text(pg, 514, yt, f"{T['pkg']:,}", 8.5, right=True, font="th")
    yb = 550
    text(pg, 53, yb, "จัดส่งข้อมูลให้ท่านทางระบบ......................", 8.5, font="th")
    text(pg, 53, yb + 20, "1.  EDI", 8.5, font="th"); text(pg, 53, yb + 40, "2.  E-MAIL", 8.5, font="th")
    for i, t in enumerate(("MISS RATTANA", "Import Documentation", "TEL:   02-059 6253 (Direct)")):
        text(pg, 333, yb + i * 20, t, 9, font="th")
        pg.insert_text((333, yb + i * 20 + 8), "." * 70, fontsize=5, fontname="th", fontfile=FONTFILES["th"])
    doc.subset_fonts(); doc.save(path, garbage=4, deflate=True)


def main():
    mf, out = find_mf(sys.argv[1] if len(sys.argv) > 1 else ".")
    if len(sys.argv) > 2: out = sys.argv[2]
    os.makedirs(out, exist_ok=True)
    vessel, bls, GT = read_mf(mf)
    if not bls: sys.exit(f"!! ไม่พบ B/L ในไฟล์ {mf}")
    if not any(b["shed"] for b in bls.values()):
        sys.exit("!! ไฟล์นี้ไม่มีช่อง 1) SHED NO. / 2) ประเทศ / 3) CARGO MOVEMENT  ใช้ MF ที่ export แบบมีข้อมูลนี้")
    summ, nosize = summarize(bls)
    T = dict(bls=sum(x["bls"] for x in summ), cntrs=sum(x["cntrs"] for x in summ),
             gw=round(sum(x["gw"] for x in summ), 3), pkg=sum(x["pkg"] for x in summ))
    vessel_name = re.sub(r"\s+(\S+)$", r" V.\1", vessel)
    write_summary(os.path.join(out, "SUMMARY.pdf"), vessel_name, summ)
    write_sum_edi(os.path.join(out, "SUM-EDI.pdf"), vessel_name, summ, T)

    print(f"MF: {mf}\nVESSEL: {vessel_name}")
    for x in summ:
        print(f"{LABEL.get(x['g'], x['g']):12} {x['n20']:>3}x20' {x['n40']:>3}x40' BL {x['bls']:>3} {x['gw']:>14,.3f} KGS {x['pkg']:>7,} PKGS")
    print(f"TOTAL        {T['cntrs']} CNTRS  BL {T['bls']}  {T['gw']:,.3f}  {T['pkg']:,}")
    if GT:
        ok = (round(GT[0], 3), GT[1]) == (T["gw"], T["pkg"])
        print(f"MF GRAND TOTAL {GT[0]:,.3f} / {GT[1]:,} -> {'OK' if ok else 'MISMATCH'}")
    unknown = sorted({b['shed'] for b in bls.values()} - {"2835", *SHED_GROUP})
    if unknown: print("SHED อื่นที่นับรวมเป็น LCH:", ", ".join(unknown))
    if nosize: print("ตู้ที่ไม่ทราบขนาด:", ", ".join(nosize))
    print("SIGNATURE:", "ใส่แล้ว" if os.path.exists(SIGNATURE) else "ไม่มี assets/signature.png (ไม่ใส่ลายเซ็น)")
    print(f"OUTPUT: {os.path.join(out, 'SUMMARY.pdf')} , {os.path.join(out, 'SUM-EDI.pdf')}")


if __name__ == "__main__":
    main()
