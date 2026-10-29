"""SUMMARY: split one full MF / MANIFEST .xls (CARGO MANIFEST with the SHED NO. / country / CARGO MOVEMENT
column) into the main port (LAEM CHABANG or BANGKOK), LKB (GATE2) and the transit countries, add the SOC
containers from CNTRS.xls (containers in CNTRS that are not in the manifest), and write SUMMARY.pdf + SUM-EDI.pdf.

usage: python sum_mf.py [job folder | MF file] [output folder] [--cntrs FILE]
                        [--name "MISS RATTANA"] [--position "Import Documentation"] [--tel "02-059 6253 (Direct)"]
  job folder (default: current folder): files are looked up in <job>/input, then <job>
  output folder default: <job>/output
"""
import sys, os, re, glob, argparse
import xlrd, pymupdf

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "..", "assets")
LETTERHEAD = os.path.join(ASSETS, "letterhead.jpg")
SIGNATURE = os.path.join(ASSETS, "signature.png")  # local only, not in git

BLRE = re.compile(r"^HASL[A-Z0-9]{10,14}$")
CNRE = re.compile(r"\b([A-Z]{4}\d{7})\b")
num = lambda t: float(t.replace(",", ""))
COUNTRY = {"MM": "MYANMAR", "LA": "LAOS", "MY": "MALAYSIA", "KH": "CAMBODIA", "VN": "VIETNAM", "CN": "CHINA"}
SHED_GROUP = {"0332": "LKR"}  # any other shed = the main port
PORTS = {"LAEM CHABANG": ("LCH", "LCH"), "BANGKOK": ("BKK", "BANGKOK")}  # port of discharge -> (code, label)
LABEL = {"LCH": "LCH", "BKK": "BANGKOK", "LKR": "LKB (GATE2)", "EMPTY": "EMPTY"}
ORDER = ["LCH", "BKK", "LKR", "MYANMAR", "LAOS", "MALAYSIA"]
# SOC owner = first word of the CONSIGNEE: "ORIENT OVERSEAS CONTAINER LINE" -> "ORIENT", "NS INTERNATIONAL ..." -> "NS"
SOC_STOP = {"AGENCIES", "AGENCY", "INTERNATIONAL", "(THAILAND)", "THAILAND", "CO.", "CO.,", "CO.,LTD.", "CO.,LTD",
            "LTD.", "LTD", "LIMITED", "SHIPPING", "LOGISTICS", "CORPORATION", "CORP.", "COMPANY"}
STAFF = dict(name="MISS RATTANA", position="Import Documentation", tel="02-059 6253 (Direct)")


def xls_files(d):
    return [f for f in glob.glob(os.path.join(d, "*.xls*")) if not os.path.basename(f).startswith("~$")]


def find_files(arg, cntrs):
    if os.path.isfile(arg):
        base = os.path.dirname(os.path.abspath(arg)); mf = arg; out = os.path.join(base, "output"); dirs = [base]
    else:
        dirs = [os.path.join(arg, "input"), arg]; out = os.path.join(arg, "output"); mf = None
        for d in dirs:
            files = [f for f in xls_files(d) if not os.path.basename(f).upper().startswith("CNTRS")]
            rank = lambda f: (0 if os.path.basename(f).upper().startswith("MF") else 1 if os.path.basename(f).upper().startswith("MANIFEST") else 2, f)
            if files: mf = sorted(files, key=rank)[0]; break
        if not mf: sys.exit(f"!! No MF / MANIFEST file (.xls/.xlsx) in {arg}\\input or {arg}")
    if cntrs is None:
        for d in dirs:
            c = [f for f in xls_files(d) if os.path.basename(f).upper().startswith("CNTRS")]
            if c: cntrs = sorted(c)[0]; break
    return mf, cntrs, out


def read_mf(path):
    s = xlrd.open_workbook(path, logfile=open(os.devnull, "w")).sheet_by_index(0)
    bls, cur, vessel, port, gt = {}, None, "", "", None
    for r in range(s.nrows):
        v = [str(x).strip() for x in s.row_values(r)] + [""] * 23
        line = " ".join(x for x in v if x)
        if v[0] == "VESSEL & VOYAGE": vessel = v[2]
        if v[0] == "PORT OF LOADING" and not port: port = v[8]
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
        b["cntrs"] += CNRE.findall(v[5] + " " + v[6])   # CONTAINER NO. is in column F or G
        m = re.match(r"^(\d{2}[A-Z]\d)X\d+", v[10])
        if m: b["types"].append(m.group(1))
    total = None
    if gt is not None:
        blk = " ".join(" ".join(str(x) for x in s.row_values(r)) for r in range(gt, min(gt + 4, s.nrows)))
        m = re.search(r"GROSSWEIGHT:\s*([\d,.]+).*?TOTAL PACK:\s*([\d,]+)", blk)
        if m: total = (num(m.group(1)), int(num(m.group(2))))
    return vessel, port, bls, total


def read_cntrs(path):
    """CNTRS.xls container list -> {container: (size "20'"/"40'", consignee, empty)}"""
    s = xlrd.open_workbook(path, logfile=open(os.devnull, "w")).sheet_by_index(0)
    out = {}
    for r in range(s.nrows):
        v = [str(x).strip() for x in s.row_values(r)] + [""] * 31
        if not re.fullmatch(r"[A-Z]{4}\d{7}", v[0]): continue
        m = re.search(r"\((\d\d)'\)", " ".join(v[3:9]))
        size = (m.group(1) if m else "") or ("20" if re.search(r"\b2\d[A-Z]\d\b", " ".join(v[3:9])) else "40")
        status = next((x for x in v[15:23] if re.match(r"\d=", x)), "")
        out.setdefault(v[0], ("20'" if size == "20" else "40'", v[9], bool(re.match(r"4=", status) or "MTY" in status)))
    return out


def soc_owner(consignee):
    words = []
    for w in consignee.upper().split():
        if w in SOC_STOP: break
        words.append(w)
    return words[0] if words else (consignee.split() or ["SOC"])[0]


def group_of(b, main):
    if b["transit"] and b["country"] not in ("", "TH"): return COUNTRY.get(b["country"], b["country"])
    return SHED_GROUP.get(b["shed"], main)


def summarize(bls, main, cn_list):
    size = {c: s for c, (s, _, _) in cn_list.items()}
    for b in bls.values():
        for c in b["cntrs"]:
            if b["types"]: size.setdefault(c, "20'" if b["types"][0][0] == "2" else "40'")
    groups = {}
    for no, b in bls.items(): groups.setdefault(group_of(b, main), []).append(no)
    order = ORDER + [g for g in groups if g not in ORDER]
    summ = []
    for g in [g for g in order if g in groups]:
        nos = groups[g]
        cn = {c for n in nos for c in bls[n]["cntrs"]}
        summ.append(dict(g=g, nos=nos, bls=len(nos), cntrs=len(cn),
                         n20=sum(size.get(c) == "20'" for c in cn), n40=sum(size.get(c) == "40'" for c in cn),
                         gw=round(sum(bls[n]["gw"] for n in nos), 3), pkg=sum(bls[n]["pkg"] for n in nos)))
    in_mf = {c for b in bls.values() for c in b["cntrs"]}
    soc = {}
    # containers in CNTRS but not in the manifest:
    #   HEUNG A, not empty -> counted in the main port row; HEUNG A empty (4=MTY) -> EMPTY row in the table;
    #   any other owner -> ***SOC*** lines under TOTAL
    extra, empty_l = [], []
    for c, (sz, cons, empty) in cn_list.items():
        if c in in_mf: continue
        if not HALRE.search(cons): soc.setdefault((soc_owner(cons), sz), []).append(c)
        else: (empty_l if empty else extra).append(c)
    count = lambda cs: dict(cntrs=len(cs), n20=sum(size.get(c) == "20'" for c in cs), n40=sum(size.get(c) == "40'" for c in cs))
    if extra:
        m = next((x for x in summ if x["g"] == main), None)
        if m is None:
            m = dict(g=main, nos=[], bls=0, cntrs=0, n20=0, n40=0, gw=0.0, pkg=0); summ.insert(0, m)
        k = count(extra)
        for f in ("cntrs", "n20", "n40"): m[f] += k[f]
        m["extra"] = sorted(extra)
    if empty_l: summ.append(dict(g="EMPTY", empty=True, nos=sorted(empty_l), bls=0, gw=0.0, pkg=0, **count(empty_l)))
    soc = sorted((dict(owner=o, size=sz, cntrs=sorted(cs)) for (o, sz), cs in soc.items()),
                 key=lambda s: (s["owner"], s["size"]))
    nosize = sorted(c for c in in_mf if c not in size)
    not_in_cntrs = sorted(in_mf - set(cn_list)) if cn_list else []
    return summ, soc, nosize, not_in_cntrs


# ---------------------------------------------------------------- drawing
TIMES, TIMESB = r"C:\Windows\Fonts\times.ttf", r"C:\Windows\Fonts\timesbd.ttf"
FONTFILES = {"tr": TIMES, "tb": TIMESB, "th": r"C:\Windows\Fonts\tahoma.ttf", "thb": r"C:\Windows\Fonts\tahomabd.ttf"}
HALRE = re.compile(r"HEUNG[\s-]*A\b", re.I)
AGENT1 = "HEUNG A LINE CO., LTD. C/O HEUNG A LINE (THAILAND) CO.,LTD."
AGENT2 = "HEUNG A LINE CO.,LTD C/O HEUNG-A SHIPPING (THAILAND) CO.,LTD."


def text(pg, x, y, s, size=8, bold=False, right=False, underline=False, font=None, maxw=0):
    """maxw = widest the text may be; longer text is drawn smaller so it stays inside the page"""
    fn = font or ("tb" if bold else "tr"); ff = FONTFILES[fn]
    w = pymupdf.Font(fontfile=ff).text_length(s, size)
    if maxw and w > maxw: size *= maxw / w; w = maxw
    if right: x -= w
    pg.insert_text((x, y), s, fontsize=size, fontname=fn, fontfile=ff)
    if underline: pg.draw_line((x, y + 1.5), (x + w, y + 1.5), width=0.5)
    return w


def summary_page(doc, vessel_name, port, agent, left, place, gw, pkg, double=False):
    pg = doc.new_page(width=842, height=595)
    title = f"SUMMARY OF INWARD CARGO MANIFEST DISCHARGING AT {port}"
    text(pg, 421 - pymupdf.Font(fontfile=TIMESB).text_length(title, 8) / 2, 60, title, 8, bold=True)
    text(pg, 133, 99, "FOR MV.", 8, underline=True); text(pg, 177, 99, vessel_name, 8, underline=True, maxw=600)
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


def write_summary(path, vessel_name, port, main, summ):
    S = {x["g"]: x for x in summ}
    doc = pymupdf.open()
    if main in S: summary_page(doc, vessel_name, port, AGENT1, port, "", S[main]["gw"], S[main]["pkg"], double=True)
    for g in [g for g in S if g not in (main, "LKR")]:
        summary_page(doc, vessel_name, port, AGENT2, "*CARGO INTRANSIT TO",
                     f"{g}{'  ' if g == 'LAOS' else ''}   P.D.R.  VIA  {port}", S[g]["gw"], S[g]["pkg"])
    if "LKR" in S: summary_page(doc, vessel_name, port, AGENT1, f"{port} /", "LKB (ESCO GATE 2)", S["LKR"]["gw"], S["LKR"]["pkg"], double=True)
    doc.subset_fonts(); doc.save(path, garbage=4, deflate=True)


def write_sum_edi(path, vessel_name, summ, T, soc, staff):
    doc = pymupdf.open(); pg = doc.new_page(width=595, height=842)
    pg.insert_image(pymupdf.Rect(45, 52, 45 + 505, 52 + 102), filename=LETTERHEAD)
    for i, (k, v) in enumerate((("AGENT  NAME  :", AGENT1), ("TAX ID:", "0993000456955"), ("VESSEL  :", vessel_name))):
        y = 194 + i * 21
        text(pg, 53, y, k, 9.5, font="thb"); text(pg, 147, y, v, 9.5, font="thb", maxw=400)
        pg.draw_line((140, y + 5), (510, y + 5), width=0.5, color=(0.5, 0.5, 0.5))
    X = [140, 263, 330, 420, 520]
    top, hh, rh = 270, 50, 20
    nrows = max(len(summ) + 1, 4)                    # always at least one blank row, at least 4 rows
    body = top + hh
    bottom = body + rh * nrows + 22
    pg.draw_rect(pymupdf.Rect(X[0], top, X[-1], body), width=0.7)
    for x in X: pg.draw_line((x, top), (x, bottom), width=0.7)
    for i in range(len(summ), nrows):                # ruled blank rows
        pg.draw_line((X[0], body + rh * i + 2), (X[-1], body + rh * i + 2), width=0.5)

    def ctext(x0, x1, y, s, size=8.5, font="thb"):
        w = pymupdf.Font(fontfile=FONTFILES[font]).text_length(s, size); text(pg, (x0 + x1 - w) / 2, y, s, size, font=font)
    ctext(X[0], X[1], top + 29, "NO. OF CONTAINERS"); ctext(X[1], X[2], top + 23, "NO. OF B/L"); ctext(X[1], X[2], top + 36, "(D/O)")
    ctext(X[2], X[3], top + 23, "TOTAL GROSS"); ctext(X[2], X[3], top + 36, "WEIGHT (KGS)")
    ctext(X[3], X[4], top + 23, "TOTAL NO. OF"); ctext(X[3], X[4], top + 36, "PACKAGES")
    y = body + 15
    for x in summ:
        text(pg, X[0] - 2, y, LABEL.get(x["g"], x["g"]), 8.5, right=True, font="thb", maxw=90)
        if x["n20"]: text(pg, 185, y, f"{x['n20']}  X 20'", 8.5, right=True, font="th")
        if x["n40"]: text(pg, 258, y, f"{x['n40']}  X 40'", 8.5, right=True, font="th")
        if not x.get("empty"):
            text(pg, 300, y, f"{x['bls']}", 8.5, right=True, font="th")
            text(pg, 412, y, f"{x['gw']:,.3f}", 8.5, right=True, font="th")
            text(pg, 514, y, f"{x['pkg']:,}", 8.5, right=True, font="th")
        y += rh
    yt = body + rh * nrows + 15
    pg.draw_line((X[0], yt - 15), (X[-1], yt - 15), width=0.7); pg.draw_line((X[0], yt + 7), (X[-1], yt + 7), width=0.7)
    text(pg, X[0] - 2, yt, "TOTAL", 8.5, right=True, font="thb")
    text(pg, 205, yt, f"{T['cntrs']}", 8.5, right=True, font="th"); text(pg, 215, yt, "CNTRS", 8.5, font="th")
    text(pg, 300, yt, f"{T['bls']}", 8.5, right=True, font="th")
    text(pg, 412, yt, f"{T['gw']:,.3f}", 8.5, right=True, font="th")
    text(pg, 514, yt, f"{T['pkg']:,}", 8.5, right=True, font="th")
    yb = yt + 30
    if soc:                                          # SOC containers: other owners, not in the manifest
        text(pg, 53, yb, "***SOC***", 8, font="thb")
        yb += 20
        for i, s in enumerate(soc):
            ys = yb + i * 16
            text(pg, 118, ys, s["owner"], 8, right=True, font="thb", maxw=70)
            text(pg, 128, ys, f"{s['size']} :", 8, font="thb")
            text(pg, 205, ys, f"{len(s['cntrs'])}", 8, right=True, font="thb"); text(pg, 210, ys, "CNTRS", 8, font="thb")
            text(pg, 300, ys, "[SUM-EDI, MANIFEST AS PER ATTACHED]", 8, font="thb")
        yg = yb + len(soc) * 16 + 10
        text(pg, 45, yg, "GRAND TOTAL", 8, font="thb", underline=True)
        text(pg, 205, yg, f"{T['cntrs'] + sum(len(s['cntrs']) for s in soc)}", 8, right=True, font="thb", underline=True)
        text(pg, 210, yg, "CNTRS", 8, font="thb", underline=True)
        yb = yg + 30
    yb = max(yb, 550)
    text(pg, 120, yb, "DATA SUBMITTED TO THE PORT BY ......................", 8.5, font="th")
    text(pg, 245, yb + 20, "1.  EDI", 8.5, font="th"); text(pg, 245, yb + 38, "2.  E-MAIL", 8.5, font="th")
    for i, t in enumerate((staff["name"], staff["position"], f"TEL:   {staff['tel']}")):
        if not t.strip() or t.strip() == "TEL:": continue
        y = yb + 70 + i * 20
        text(pg, 390, y, t, 9, font="tr", maxw=160)
        pg.draw_line((385, y + 5), (545, y + 5), width=0.5)
    doc.subset_fonts(); doc.save(path, garbage=4, deflate=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("job", nargs="?", default=".")
    ap.add_argument("out", nargs="?")
    ap.add_argument("--cntrs")
    for k, v in STAFF.items(): ap.add_argument("--" + k, default=v)
    a = ap.parse_args()
    mf, cntrs, out = find_files(a.job, a.cntrs)
    out = a.out or out
    os.makedirs(out, exist_ok=True)
    vessel, port, bls, GT = read_mf(mf)
    if not bls: sys.exit(f"!! No B/L (HASL...) in {mf}")
    if not any(b["shed"] for b in bls.values()):
        sys.exit("!! This file has no 1) SHED NO. / 2) country / 3) CARGO MOVEMENT column - use the full MF export")
    port = port or "LAEM CHABANG"
    main_code = PORTS.get(port, (port, port))[0]
    LABEL.setdefault(main_code, PORTS.get(port, (port, port))[1])
    cn_list = read_cntrs(cntrs) if cntrs else {}
    summ, soc, nosize, not_in_cntrs = summarize(bls, main_code, cn_list)
    T = dict(bls=sum(x["bls"] for x in summ), cntrs=sum(x["cntrs"] for x in summ),
             gw=round(sum(x["gw"] for x in summ), 3), pkg=sum(x["pkg"] for x in summ))
    vessel_name = re.sub(r"\s+(\S+)$", r" V.\1", vessel)
    staff = dict(name=a.name, position=a.position, tel=a.tel)
    write_summary(os.path.join(out, "SUMMARY.pdf"), vessel_name, port, main_code, [x for x in summ if not x.get("empty")])
    write_sum_edi(os.path.join(out, "SUM-EDI.pdf"), vessel_name, summ, T, soc, staff)

    print(f"MF: {mf}\nCNTRS: {cntrs or '- (no SOC check)'}\nVESSEL: {vessel_name}   PORT: {port}")
    for x in summ:
        print(f"{LABEL.get(x['g'], x['g']):12} {x['n20']:>3}x20' {x['n40']:>3}x40' BL {x['bls']:>3} {x['gw']:>14,.3f} KGS {x['pkg']:>7,} PKGS")
        if x.get("extra"): print(f"             + HEUNG A not in MF: {', '.join(x['extra'])}")
    print(f"TOTAL        {T['cntrs']} CNTRS  BL {T['bls']}  {T['gw']:,.3f}  {T['pkg']:,}")
    if soc: print("***SOC***")
    for s in soc: print(f"{s['owner']:12} {s['size']} : {len(s['cntrs'])} CNTRS")
    if soc: print(f"GRAND TOTAL  {T['cntrs'] + sum(len(s['cntrs']) for s in soc)} CNTRS")
    if GT:
        ok = (round(GT[0], 3), GT[1]) == (T["gw"], T["pkg"])
        print(f"MF GRAND TOTAL {GT[0]:,.3f} / {GT[1]:,} -> {'OK' if ok else 'MISMATCH'}")
    if not_in_cntrs: print("IN MANIFEST BUT NOT IN CNTRS:", ", ".join(not_in_cntrs))
    if nosize: print("CONTAINERS WITHOUT SIZE:", ", ".join(nosize))
    print("SIGNATURE:", "yes" if os.path.exists(SIGNATURE) else "none (assets/signature.png missing)")
    print(f"OUTPUT: {os.path.join(out, 'SUMMARY.pdf')} , {os.path.join(out, 'SUM-EDI.pdf')}")


if __name__ == "__main__":
    main()
