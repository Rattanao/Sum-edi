"""SUM-EDI: split one full MF / MANIFEST .xls (CARGO MANIFEST with the SHED NO. / country / CARGO MOVEMENT
column) by DELIVERY (THLCH, THLKR, THBMT, THSCT, THBKK from the CNTRS container list, the transit countries,
TRANSHIPMENT, EMPTY), add the SOC containers from CNTRS (containers in CNTRS that are not in the manifest),
and write SUMMARY.pdf + SUM-EDI.pdf.

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
# DELIVERY: "Port Of Delivery" of the CNTRS page; when that is the port of discharge, the shed tells the terminal
SHED_DELIVERY = {"0332": "THLKR", "0110": "THBMT", "0302": "THSCT", "0122": "THBKK", "0145": "THBKK"}
PORTS = {"LAEM CHABANG": "THLCH", "BANGKOK": "THBKK"}  # port of discharge -> delivery code
ORDER = ["THLCH", "THLKR", "THBMT", "THSCT", "THBKK", "MYANMAR", "LAOS", "MALAYSIA", "CHINA", "CAMBODIA",
         "TRANSHIPMENT", "EMPTY"]
TRANSIT = {"MYANMAR", "LAOS", "MALAYSIA", "CHINA", "CAMBODIA", "VIETNAM"}
# SOC owner = first word of the CONSIGNEE: "ORIENT OVERSEAS CONTAINER LINE" -> "ORIENT", "NS INTERNATIONAL ..." -> "NS"
SOC_STOP = {"AGENCIES", "AGENCY", "INTERNATIONAL", "(THAILAND)", "THAILAND", "CO.", "CO.,", "CO.,LTD.", "CO.,LTD",
            "LTD.", "LTD", "LIMITED", "SHIPPING", "LOGISTICS", "CORPORATION", "CORP.", "COMPANY"}
STAFF = dict(name="MISS RATTANA", position="Import Documentation", tel="02-059 6253 (Direct)")


def xls_files(d):
    return [f for f in glob.glob(os.path.join(d, "*.xls*")) if not os.path.basename(f).startswith("~$")]


def is_cntrs(f):
    """the CNTRS container list (any file name, e.g. CNTRS.xls or INWARD.xls) says CONTAINER LIST FOR FEEDER"""
    try:
        s = xlrd.open_workbook(f, logfile=open(os.devnull, "w")).sheet_by_index(0)
        return any("CONTAINER LIST" in " ".join(str(x) for x in s.row_values(r)) for r in range(min(s.nrows, 15)))
    except Exception:
        return os.path.basename(f).upper().startswith("CNTRS")


def find_files(arg, cntrs):
    if os.path.isfile(arg):
        base = os.path.dirname(os.path.abspath(arg)); mf = arg; out = os.path.join(base, "output"); dirs = [base]
    else:
        dirs = [os.path.join(arg, "input"), arg]; out = os.path.join(arg, "output"); mf = None
        for d in dirs:
            files = [f for f in xls_files(d) if not is_cntrs(f)]
            rank = lambda f: (0 if os.path.basename(f).upper().startswith("MF") else 1 if os.path.basename(f).upper().startswith("MANIFEST") else 2, f)
            if files: mf = sorted(files, key=rank)[0]; break
        if not mf: sys.exit(f"!! No MF / MANIFEST file (.xls/.xlsx) in {arg}\\input or {arg}")
    if cntrs is None:
        for d in dirs:
            c = [f for f in xls_files(d) if is_cntrs(f)]
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
            bls[cur] = dict(pkg=0, gw=0.0, shed="", country="", transit=False, tranship=False, cntrs=[], types=[])
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
        if "CARGO MOVEMENT" in v[15] and "TRANSHIP" in v[15]: b["tranship"] = True
        b["cntrs"] += CNRE.findall(v[5] + " " + v[6])   # CONTAINER NO. is in column F or G
        m = re.match(r"^(\d{2}[A-Z]\d)X\d+", v[10])
        if m: b["types"].append(m.group(1))
    total = None
    if gt is not None:
        blk = " ".join(" ".join(str(x) for x in s.row_values(r)) for r in range(gt, min(gt + 4, s.nrows)))
        m = re.search(r"GROSSWEIGHT:\s*([\d,.]+).*?TOTAL PACK:\s*([\d,]+)", blk)
        if m: total = (num(m.group(1)), int(num(m.group(2))))
    return vessel, port, bls, total


def read_cntrs(path, discharge):
    """CNTRS container list -> {container: (size "20'"/"40'", consignee, empty, delivery)}"""
    s = xlrd.open_workbook(path, logfile=open(os.devnull, "w")).sheet_by_index(0)
    out, deliv, shed = {}, discharge, ""
    for r in range(s.nrows):
        v = [str(x).strip() for x in s.row_values(r)] + [""] * 31
        if "Port Of Delivery" in v:
            deliv = next((x for x in v[v.index("Port Of Delivery") + 1:] if x), discharge).upper()
        if v[0].startswith("SHED NUMBER"):
            m = re.match(r"\d+", next((x for x in v[1:] if x), ""))
            if m: shed = m.group(0).zfill(4)
        if not re.fullmatch(r"[A-Z]{4}\d{7}", v[0]): continue
        m = re.search(r"\((\d\d)'\)", " ".join(v[3:9]))
        size = (m.group(1) if m else "") or ("20" if re.search(r"\b2\d[A-Z]\d\b", " ".join(v[3:9])) else "40")
        status = next((x for x in v[15:23] if re.match(r"\d=", x)), "")
        d = SHED_DELIVERY.get(shed, deliv) if deliv == discharge else deliv
        out.setdefault(v[0], ("20'" if size == "20" else "40'", v[9], bool(re.match(r"4=", status) or "MTY" in status), d))
    return out


def soc_owner(consignee):
    words = []
    for w in consignee.upper().split():
        if w in SOC_STOP: break
        words.append(w)
    return words[0] if words else (consignee.split() or ["SOC"])[0]


def group_of(b, main, cn_list):
    """DELIVERY row of a B/L: transit country, TRANSHIPMENT, else the delivery of its containers in CNTRS"""
    if b["transit"] and b["country"] not in ("", "TH"): return COUNTRY.get(b["country"], b["country"])
    if b["tranship"]: return "TRANSHIPMENT"
    for c in b["cntrs"]:
        if c in cn_list: return cn_list[c][3]
    return SHED_DELIVERY.get(b["shed"], main)


def summarize(bls, main, cn_list):
    size = {c: s for c, (s, _, _, _) in cn_list.items()}
    for b in bls.values():
        for c in b["cntrs"]:
            if b["types"]: size.setdefault(c, "20'" if b["types"][0][0] == "2" else "40'")
    groups = {}
    for no, b in bls.items(): groups.setdefault(group_of(b, main, cn_list), []).append(no)
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
    #   HEUNG A, not empty -> counted in its DELIVERY row; HEUNG A empty (4=MTY) -> EMPTY row in the table;
    #   any other owner -> ***SOC*** lines under GRAND TOTAL
    extra, empty_l = {}, []
    for c, (sz, cons, empty, d) in cn_list.items():
        if c in in_mf: continue
        if not HALRE.search(cons): soc.setdefault((soc_owner(cons), sz), []).append(c)
        elif empty: empty_l.append(c)
        else: extra.setdefault(d, []).append(c)
    count = lambda cs: dict(cntrs=len(cs), n20=sum(size.get(c) == "20'" for c in cs), n40=sum(size.get(c) == "40'" for c in cs))
    for d, cs in extra.items():
        m = next((x for x in summ if x["g"] == d), None)
        if m is None:
            m = dict(g=d, nos=[], bls=0, cntrs=0, n20=0, n40=0, gw=0.0, pkg=0); summ.append(m)
        k = count(cs)
        for f in ("cntrs", "n20", "n40"): m[f] += k[f]
        m["extra"] = sorted(cs)
    rank = {g: i for i, g in enumerate(ORDER)}
    summ.sort(key=lambda x: rank.get(x["g"], len(ORDER) - 1))
    if empty_l: summ.append(dict(g="EMPTY", empty=True, nos=sorted(empty_l), bls=0, gw=0.0, pkg=0, **count(empty_l)))
    soc = sorted((dict(owner=o, size=sz, cntrs=sorted(cs)) for (o, sz), cs in soc.items()),
                 key=lambda s: (s["owner"], s["size"]))
    nosize = sorted(c for c in in_mf if c not in size)
    not_in_cntrs = sorted(in_mf - set(cn_list)) if cn_list else []
    return summ, soc, nosize, not_in_cntrs


# ---------------------------------------------------------------- drawing
TIMES, TIMESB = r"C:\Windows\Fonts\times.ttf", r"C:\Windows\Fonts\timesbd.ttf"
FONTFILES = {"tr": TIMES, "tb": TIMESB, "th": r"C:\Windows\Fonts\tahoma.ttf", "thb": r"C:\Windows\Fonts\tahomabd.ttf",
             "ab": r"C:\Windows\Fonts\arialbd.ttf"}
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
    """one page for the port (every delivery in Thailand except THLKR), one per transit country, one for THLKR"""
    S = {x["g"]: x for x in summ if x["bls"]}
    port_g = [g for g in S if g not in TRANSIT and g != "THLKR"]
    doc = pymupdf.open()
    if port_g:
        summary_page(doc, vessel_name, port, AGENT1, port, "", round(sum(S[g]["gw"] for g in port_g), 3),
                     sum(S[g]["pkg"] for g in port_g), double=True)
    for g in [g for g in S if g in TRANSIT]:
        summary_page(doc, vessel_name, port, AGENT2, "*CARGO INTRANSIT TO",
                     f"{g}{'  ' if g == 'LAOS' else ''}   P.D.R.  VIA  {port}", S[g]["gw"], S[g]["pkg"])
    if "THLKR" in S: summary_page(doc, vessel_name, port, AGENT1, f"{port} /", "LKB (ESCO GATE 2)", S["THLKR"]["gw"], S["THLKR"]["pkg"], double=True)
    doc.subset_fonts(); doc.save(path, garbage=4, deflate=True)


def write_sum_edi(path, vessel, port, summ, T, soc, staff):
    doc = pymupdf.open(); pg = doc.new_page(width=595, height=842)
    pg.insert_image(pymupdf.Rect(45, 52, 45 + 505, 52 + 102), filename=LETTERHEAD)
    for i, (k, v) in enumerate((("AGENT  NAME  :", AGENT1), ("TAX ID:", "0993000456955"), ("VESSEL  :", vessel),
                                ("PORT OF DISCHARGE :", port))):
        y = 194 + i * 21
        text(pg, 54, y, k, 9.5, font="ab")
        x = 185 if i < 3 else 54 + pymupdf.Font(fontfile=FONTFILES["ab"]).text_length(k, 9.5) + 3
        text(pg, x, y, v, 9.5, font="ab", maxw=541 - x)
        pg.draw_line((185 if i < 3 else 54, y + 5), (541, y + 5), width=0.5)
    # only the DELIVERY rows that have containers, then one blank row and GRAND TOTAL
    X = [51, 147, 276, 347, 445, 541]
    top, hh, rh = 279, 43, 20
    nrows = len(summ) + 1
    body = top + hh
    bottom = body + rh * nrows + 22
    pg.draw_rect(pymupdf.Rect(X[0], top, X[-1], bottom), width=0.8)
    for x in X[1:-1]: pg.draw_line((x, top), (x, bottom), width=0.6)
    for i in range(nrows + 1): pg.draw_line((X[0], body + rh * i), (X[-1], body + rh * i), width=0.6)
    pg.draw_line((X[0], bottom - 2), (X[-1], bottom - 2), width=0.6)          # double line under GRAND TOTAL

    def ctext(x0, x1, y, s, size=8.5, font="tb"):
        w = pymupdf.Font(fontfile=FONTFILES[font]).text_length(s, size); text(pg, (x0 + x1 - w) / 2, y, s, size, font=font)
    text(pg, X[0] + 3, top + 35, "DELIVERY", 8.5, font="tb")
    ctext(X[1], X[2], top + 35, "TOTAL CNTRS"); ctext(X[2], X[3], top + 35, "TOTAL B/L")
    ctext(X[3], X[4], top + 14, "TOTAL", 9.5); ctext(X[3], X[4], top + 29, "GROSS WEIGHT", 9.5)
    ctext(X[4], X[5], top + 35, "TOTAL PKGS", 9.5)
    y = body + 14
    for x in summ:
        text(pg, X[0] + 3, y, x["g"], 9, font="tb", maxw=X[1] - X[0] - 6)
        if x["n20"]: text(pg, 208, y, f"{x['n20']}   X 20'", 8.5, right=True, font="tb")
        if x["n40"]: text(pg, 268, y, f"{x['n40']}   X 40'", 8.5, right=True, font="tb")
        if x["bls"]:
            text(pg, X[3] - 6, y, f"{x['bls']}", 8.5, right=True, font="tb")
            text(pg, X[4] - 6, y, f"{x['gw']:,.3f}", 8.5, right=True, font="tb")
            text(pg, X[5] - 6, y, f"{x['pkg']:,}", 8.5, right=True, font="tb")
        y += rh
    yt = body + rh * nrows + 14
    dash = lambda n, f="{:,}": f.format(n) if n else "-"
    ctext(X[0], X[1], yt, "GRAND TOTAL", 9)
    text(pg, 220, yt, f"{T['cntrs']}", 9, right=True, font="tb"); text(pg, 234, yt, "CNTRS", 9, font="tb")
    ctext(X[2], X[3], yt, dash(T["bls"]), 9)
    text(pg, X[4] - 6, yt, dash(T["gw"], "{:,.3f}"), 9, right=True, font="tb")
    text(pg, X[5] - 6, yt, f"{T['pkg']:,}", 9, right=True, font="tb")
    yb = bottom + 34
    if soc:                                          # SOC containers: other owners, not in the manifest
        w = text(pg, 54, yb, "***", 7.5, font="ab"); w += text(pg, 54 + w, yb, "SOC", 7.5, font="ab", underline=True)
        text(pg, 54 + w, yb, "***", 7.5, font="ab")
        yb += 20
        for i, s in enumerate(soc):
            ys = yb + i * 19
            text(pg, 54, ys, s["owner"], 7.5, font="ab", maxw=100)
            text(pg, 160, ys, s["size"], 7.5, font="ab")
            text(pg, 212, ys, f"{len(s['cntrs'])}", 7.5, right=True, font="ab"); text(pg, 216, ys, "CNTRS", 7.5, font="ab")
            text(pg, 279, ys, "[SUM-EDI, MANIFEST AS PER ATTACHED]", 7.5, font="ab")
        yg = yb + len(soc) * 19 + 2
        text(pg, 54, yg, "GRAND TOTAL", 7.5, font="ab", underline=True)
        text(pg, 212, yg, f"{T['cntrs'] + sum(len(s['cntrs']) for s in soc)}", 7.5, right=True, font="ab", underline=True)
        text(pg, 216, yg, "CNTRS", 7.5, font="ab", underline=True)
        yb = yg + 20
    yb = max(yb + 14, 700)
    text(pg, 54, yb, "DATA SUBMITTED TO THE PORT BY ........................", 9, font="tr")
    text(pg, 54, yb + 19, "1.  EDI", 9, font="tr"); text(pg, 54, yb + 38, "2.  E-MAIL", 9, font="tr")
    for i, t in enumerate((staff["name"], staff["position"], f"TEL:   {staff['tel']}")):
        if not t.strip() or t.strip() == "TEL:": continue
        y = yb + i * 19
        text(pg, 347, y, t, 9, font="tr", maxw=190)
        pg.draw_line((345, y + 5), (530, y + 5), width=0.5, dashes="[2 1] 0")
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
    main_code = PORTS.get(port, "TH" + port[:3])
    cn_list = read_cntrs(cntrs, main_code) if cntrs else {}
    summ, soc, nosize, not_in_cntrs = summarize(bls, main_code, cn_list)
    summ = [x for x in summ if x["cntrs"] or x["bls"]]
    T = dict(bls=sum(x["bls"] for x in summ), cntrs=sum(x["cntrs"] for x in summ),
             gw=round(sum(x["gw"] for x in summ), 3), pkg=sum(x["pkg"] for x in summ))
    vessel_name = re.sub(r"\s+(\S+)$", r" V.\1", vessel)
    staff = dict(name=a.name, position=a.position, tel=a.tel)
    write_summary(os.path.join(out, "SUMMARY.pdf"), vessel_name, port, main_code, [x for x in summ if not x.get("empty")])
    write_sum_edi(os.path.join(out, "SUM-EDI.pdf"), vessel, port, summ, T, soc, staff)

    print(f"MF: {mf}\nCNTRS: {cntrs or '- (no SOC check)'}\nVESSEL: {vessel}   PORT: {port}")
    for x in summ:
        print(f"{x['g']:13} {x['n20']:>3}x20' {x['n40']:>3}x40' BL {x['bls']:>3} {x['gw']:>14,.3f} KGS {x['pkg']:>7,} PKGS")
        if x.get("extra"): print(f"              + HEUNG A not in MF: {', '.join(x['extra'])}")
    print(f"GRAND TOTAL   {T['cntrs']} CNTRS  BL {T['bls']}  {T['gw']:,.3f}  {T['pkg']:,}")
    if soc: print("***SOC***")
    for s in soc: print(f"{s['owner']:13} {s['size']} : {len(s['cntrs'])} CNTRS")
    if soc: print(f"GRAND TOTAL   {T['cntrs'] + sum(len(s['cntrs']) for s in soc)} CNTRS")
    if GT:
        ok = (round(GT[0], 3), GT[1]) == (T["gw"], T["pkg"])
        print(f"MF GRAND TOTAL {GT[0]:,.3f} / {GT[1]:,} -> {'OK' if ok else 'MISMATCH'}")
    if not_in_cntrs: print("IN MANIFEST BUT NOT IN CNTRS:", ", ".join(not_in_cntrs))
    if nosize: print("CONTAINERS WITHOUT SIZE:", ", ".join(nosize))
    print("SIGNATURE:", "yes" if os.path.exists(SIGNATURE) else "none (assets/signature.png missing)")
    print(f"OUTPUT: {os.path.join(out, 'SUMMARY.pdf')} , {os.path.join(out, 'SUM-EDI.pdf')}")


if __name__ == "__main__":
    main()
