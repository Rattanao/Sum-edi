"""SUM-EDI: split one full MF / MANIFEST .xls or .pdf (CARGO MANIFEST with the SHED NO. / country / CARGO MOVEMENT
column) by DELIVERY (THLCH, THLKR, THBMT, THSCT, THBKK from the CNTRS container list, the transit countries,
TRANSHIPMENT, EMPTY), add the SOC containers from CNTRS (containers in CNTRS that are not in the manifest),
and write SUMMARY.pdf + SUM-EDI.pdf.

usage: python sum_mf.py [job folder | MF file] [output folder] [--cntrs FILE] [--gate 8|9|text]
                        [--name "MISS RATTANA"] [--position "Import Documentation"] [--tel "02-059 6253 (Direct)"]
  job folder (default: current folder): files are looked up in <job>/input, then <job>
  MF and CNTRS may each be Excel (.xls/.xlsx) or PDF, and may be split over several files
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
TRANSITRE = re.compile(r"TRANSIT\s+TO\s+(" + "|".join(sorted(TRANSIT)) + r")\b")
# SUMMARY page "<PORT> / <place>"; THLKR = LATKRABANG + the gate (--gate)
PLACE = {"THLKR": "LATKRABANG (GATE 2)", "THBMT": "BANGKOK MODERN TERMINAL", "THSCT": "SIAM CONTAINER TERMINAL",
         "THBKK": "BANGKOK"}
GATES = {"8": "(TLC GATE 8)", "9": "(ESCO GATE 9)"}
# SOC owner = first word of the CONSIGNEE: "ORIENT OVERSEAS CONTAINER LINE" -> "ORIENT", "NS INTERNATIONAL ..." -> "NS"
SOC_STOP = {"AGENCIES", "AGENCY", "INTERNATIONAL", "(THAILAND)", "THAILAND", "CO.", "CO.,", "CO.,LTD.", "CO.,LTD",
            "LTD.", "LTD", "LIMITED", "SHIPPING", "LOGISTICS", "CORPORATION", "CORP.", "COMPANY"}
STAFF = dict(name="MISS RATTANA", position="Import Documentation", tel="02-059 6253 (Direct)")


def input_files(d):
    return [f for f in glob.glob(os.path.join(d, "*")) if re.search(r"\.(xlsx?|pdf)$", f, re.I)
            and not os.path.basename(f).startswith("~$")]


def pdf_lines(path):
    """text lines of a PDF, words of one line joined left to right. The bigger page-header text is kept on its
    own line, so a header printed over a container row does not mix with it."""
    out = []
    for pg in pymupdf.open(path):
        sp = []
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for t in l["spans"]:
                    if t["text"].strip(): sp.append((t["origin"][1], t["size"] >= 11, t["origin"][0], t["text"].strip()))
        sp.sort()
        rows = []
        for y, big, x, t in sp:
            r = next((r for r in rows[-4:] if abs(r[0] - y) <= 2 and r[1] == big), None)
            if r: r[2].append((x, t))
            else: rows.append([y, big, [(x, t)]])
        out += [" ".join(t for x, t in sorted(r[2])) for r in rows]
    return out


def kind_of(f):
    """-> ("cntrs" (CONTAINER LIST FOR FEEDER) / "mf" (CARGO MANIFEST) / "", voyage e.g. 2611S) - by content, any file name"""
    try:
        if f.lower().endswith(".pdf"): head = " ".join(pdf_lines(f)[:15])
        else:
            s = xlrd.open_workbook(f, logfile=open(os.devnull, "w")).sheet_by_index(0)
            head = " ".join(" ".join(str(x) for x in s.row_values(r)) for r in range(min(s.nrows, 15)))
    except Exception:
        return "", ""
    m = re.search(r"\b(\d{4}[NSEW])\b", head)
    voy = m.group(1) if m else ""
    if "CONTAINER LIST" in head: return "cntrs", voy
    if "CARGO MANIFEST" in head or "VESSEL & VOYAGE" in head: return "mf", voy
    return "", voy


def find_files(arg, cntrs):
    """-> (MF files, CNTRS files, output folder). MF and CNTRS may each be split over several files; only the
    files of the same voyage as the MF are used (a folder may hold files of other vessels)."""
    if os.path.isfile(arg):
        base = os.path.dirname(os.path.abspath(arg)); mf = [arg]; out = os.path.join(base, "output"); dirs = [base]
    else:
        dirs = [os.path.join(arg, "input"), arg]; out = os.path.join(arg, "output"); mf = []
    cn = [cntrs] if cntrs else []
    for d in dirs:
        kinds = {f: kind_of(f) for f in sorted(input_files(d)) if f not in mf + cn}
        if not mf:
            rank = lambda f: (0 if os.path.basename(f).upper().startswith("MF") else 1 if os.path.basename(f).upper().startswith("MANIFEST") else 2, f)
            mf = sorted((f for f, k in kinds.items() if k[0] == "mf"), key=rank)
        if mf:
            voy = kind_of(mf[0])[1]
            same = lambda f: not voy or kinds.get(f, ("", ""))[1] in (voy, "")
            mf = [mf[0]] + [f for f in mf[1:] if same(f)]
            if not cn: cn = [f for f, k in kinds.items() if k[0] == "cntrs" and same(f)]
            break
    if not mf: sys.exit(f"!! No MF / MANIFEST file (.xls/.xlsx/.pdf) in {arg}\\input or {arg}")
    return mf, cn, out


def new_bl():
    return dict(pkg=0, gw=0.0, shed="", country="", transit=False, tranship=False, cntrs=[], types=[])


def read_mf_pdf(path):
    """MF / MANIFEST printed to PDF: same fields as the Excel export, read from the text lines"""
    bls, cur, vessel, port, total, in_gt = {}, None, "", "", None, False
    for l in pdf_lines(path):
        m = re.search(r"VESSEL & VOYAGE\s+(.*?)\s+NATIONALITY", l)
        if m and not vessel: vessel = m.group(1).strip()
        m = re.search(r"PORT OF DISCHARGE\s+(.*?)\s+NEXT PORT", l)
        if m and not port: port = m.group(1).strip()
        if "GRAND TOTAL" in l: in_gt, total = True, [0.0, 0]
        if in_gt:
            m = re.search(r"GROSSWEIGHT:\s*([\d,.]+)", l)
            if m: total[0] = num(m.group(1))
            m = re.search(r"TOTAL PACK:\s*([\d,]+)", l)
            if m: total[1] = int(num(m.group(1))); in_gt = False
        m = re.match(r"^(HASL[A-Z0-9]{10,14})\b", l)
        if m:
            cur = m.group(1); bls[cur] = b = new_bl()
            m = re.search(r"\b(\d[\d,]*)\s+[A-Z]{2}\s*\(", l)
            if m: b["pkg"] = int(num(m.group(1)))
        if cur is None: continue
        if re.match(r"(PortOfDischarge|VESSEL & VOYAGE|.*PORT TOTAL|.*GRAND TOTAL)", l): cur = None; continue
        b = bls[cur]
        m = re.search(r"([\d,]+\.\d+)\s*KGM", l)
        if m and not b["gw"]: b["gw"] = num(m.group(1))
        m = re.search(r"SHED NO\.\s*(\d+)", l)
        if m: b["shed"] = m.group(1)
        m = re.search(r"\b2\)[^()]*\(([A-Z]{2})\)", l)
        if m: b["country"] = m.group(1)
        if "CARGO MOVEMENT" in l and "TRANSIT" in l: b["transit"] = True
        if "CARGO MOVEMENT" in l and "TRANSHIP" in l: b["tranship"] = True
        b["cntrs"] += re.findall(r"\b\d+\.\s*([A-Z]{4}\d{7})\b", l)
        b["types"] += re.findall(r"\b(\d{2}[A-Z]\d)X\d+\b", l)
        b.setdefault("text", []).append(l)
    for b in bls.values():               # MF without the country column: a B/L whose text says TRANSIT and names a
        t = " ".join(b.pop("text", []))  # transit country goes to that country (the one named most often)
        if not b["country"] and "TRANSIT" in t:
            n = {c: len(re.findall(r"\b" + c + r"\b", t)) for c in TRANSIT - {"VIETNAM"}}
            n["LAOS"] += len(re.findall(r"\bLAO\b", t))
            if max(n.values()): b["mention"] = max(n, key=n.get)
    return vessel, port, bls, tuple(total) if total else None


def read_mf_all(paths):
    """one or more MF files (Excel or PDF) -> vessel, port, B/L, MF GRAND TOTAL (summed over the files)"""
    vessel, port, bls, total = "", "", {}, None
    for f in paths:
        v, p, b, t = read_mf_pdf(f) if f.lower().endswith(".pdf") else read_mf(f)
        vessel, port = vessel or v, port or p
        bls.update(b)
        if t: total = (round((total or (0, 0))[0] + t[0], 3), (total or (0, 0))[1] + t[1])
    return vessel, port, bls, total


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
            bls[cur] = new_bl()
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


def read_cntrs_pdf(path, discharge):
    """CNTRS container list printed to PDF -> same as read_cntrs"""
    out, deliv, shed = {}, discharge, ""
    for l in pdf_lines(path):
        m = re.search(r"Port Of Delivery\s+([A-Z]{5})", l)
        if m: deliv = m.group(1)
        m = re.match(r"SHED NUMBER:\s*(\d+)", l)
        if m: shed = m.group(1).zfill(4)
        m = re.match(r"^([A-Z]{4}\d{7})\s+\d+\s+\S+\s+(\w{4})\s*\((\d\d)'\)\s*(.*?)\s*[\d,]+\.\d+\s*KGM\s*(\S*)", l)
        if not m: continue
        d = SHED_DELIVERY.get(shed, deliv) if deliv == discharge else deliv
        t = TRANSITRE.search(l)
        out.setdefault(m.group(1), ("20'" if m.group(3) == "20" else "40'", m.group(4),
                                    m.group(5).startswith("4=") or "MTY" in m.group(5), d, t.group(1) if t else ""))
    return out


def read_cntrs_all(paths, discharge):
    out = {}
    for f in paths:
        for c, x in (read_cntrs_pdf(f, discharge) if f.lower().endswith(".pdf") else read_cntrs(f, discharge)).items():
            out.setdefault(c, x)
    return out


def read_cntrs(path, discharge):
    """CNTRS container list -> {container: (size "20'"/"40'", consignee, empty, delivery, transit country in REMARK)}"""
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
        t = TRANSITRE.search(" ".join(v[20:]))
        out.setdefault(v[0], ("20'" if size == "20" else "40'", v[9], bool(re.match(r"4=", status) or "MTY" in status), d,
                              t.group(1) if t else ""))
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
    if not b["country"]:                 # MF without the country column: "TRANSIT TO <COUNTRY>" in the CNTRS REMARK
        t = next((cn_list[c][4] for c in b["cntrs"] if c in cn_list and cn_list[c][4]), "")
        if t or b.get("mention"): return t or b["mention"]
    for c in b["cntrs"]:
        if c in cn_list: return cn_list[c][3]
    return SHED_DELIVERY.get(b["shed"], main)


def summarize(bls, main, cn_list):
    size = {c: x[0] for c, x in cn_list.items()}
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
    for c, (sz, cons, empty, d, _) in cn_list.items():
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


def gate_text(gate):
    """GATE field -> text after LATKRABANG: "8" -> (TLC GATE 8), "9" -> (ESCO GATE 9), empty -> (GATE 2)"""
    g = (gate or "").strip().upper()
    if not g: return "(GATE 2)"
    if g in GATES: return GATES[g]
    if g.isdigit(): return f"(GATE {g})"
    return g if g.startswith("(") else f"({g})"


def write_summary(path, vessel_name, port, main, summ, gate=""):
    """one page per DELIVERY row with B/L, in the SUM-EDI order"""
    PLACE["THLKR"] = f"LATKRABANG {gate_text(gate)}"
    doc = pymupdf.open()
    for x in [x for x in summ if x["bls"]]:
        g = x["g"]
        if g == main: summary_page(doc, vessel_name, port, AGENT1, port, "", x["gw"], x["pkg"], double=True)
        elif g in TRANSIT:
            summary_page(doc, vessel_name, port, AGENT2, "*CARGO INTRANSIT TO",
                         f"{g}{'  ' if g == 'LAOS' else ''}   P.D.R.  VIA  {port}", x["gw"], x["pkg"])
        else: summary_page(doc, vessel_name, port, AGENT1, f"{port} /", PLACE.get(g, g), x["gw"], x["pkg"], double=True)
    doc.subset_fonts(); doc.save(path, garbage=4, deflate=True)


def write_sum_edi(path, vessel, port, summ, T, soc, staff):
    doc = pymupdf.open(); pg = doc.new_page(width=595, height=842)
    pg.insert_image(pymupdf.Rect(45, 52, 45 + 505, 52 + 102), filename=LETTERHEAD)
    for i, (k, v) in enumerate((("AGENT  NAME  :", AGENT1), ("TAX ID:", "0993000456955"), ("VESSEL  :", vessel),
                                ("DISCHARGE :", port))):
        y = 194 + i * 21
        text(pg, 54, y, k, 9.5, font="ab")
        text(pg, 185, y, v, 9.5, font="ab", maxw=541 - 185)
        pg.draw_line((185, y + 5), (541, y + 5), width=0.5)
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
        w = text(pg, 54, yb, "***", 9.5, font="ab"); w += text(pg, 54 + w, yb, "SOC", 9.5, font="ab", underline=True)
        text(pg, 54 + w, yb, "***", 9.5, font="ab")
        yb += 24
        for i, s in enumerate(soc):
            ys = yb + i * 22
            text(pg, 54, ys, s["owner"], 9.5, font="ab", maxw=100)
            text(pg, 170, ys, s["size"], 9.5, font="ab")
            text(pg, 232, ys, f"{len(s['cntrs'])}", 9.5, right=True, font="ab"); text(pg, 237, ys, "CNTRS", 9.5, font="ab")
            text(pg, 300, ys, "[SUM-EDI, MANIFEST AS PER ATTACHED]", 9.5, font="ab")
        yg = yb + len(soc) * 22 + 4
        text(pg, 54, yg, "GRAND TOTAL", 9.5, font="ab", underline=True)
        text(pg, 232, yg, f"{T['cntrs'] + sum(len(s['cntrs']) for s in soc)}", 9.5, right=True, font="ab", underline=True)
        text(pg, 237, yg, "CNTRS", 9.5, font="ab", underline=True)
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
    ap.add_argument("--gate", default="", help='LATKRABANG gate: 8 = (TLC GATE 8), 9 = (ESCO GATE 9), or any text')
    for k, v in STAFF.items(): ap.add_argument("--" + k, default=v)
    a = ap.parse_args()
    mf, cntrs, out = find_files(a.job, a.cntrs)
    out = a.out or out
    os.makedirs(out, exist_ok=True)
    vessel, port, bls, GT = read_mf_all(mf)
    if not bls: sys.exit(f"!! No B/L (HASL...) in {', '.join(mf)}")
    if not any(b["shed"] for b in bls.values()) and not cntrs:
        sys.exit("!! The MF has no 1) SHED NO. / 2) country / 3) CARGO MOVEMENT column - give the CNTRS file too, or use the full MF export")
    port = port or "LAEM CHABANG"
    main_code = PORTS.get(port, "TH" + port[:3])
    cn_list = read_cntrs_all(cntrs, main_code)
    summ, soc, nosize, not_in_cntrs = summarize(bls, main_code, cn_list)
    summ = [x for x in summ if x["cntrs"] or x["bls"]]
    T = dict(bls=sum(x["bls"] for x in summ), cntrs=sum(x["cntrs"] for x in summ),
             gw=round(sum(x["gw"] for x in summ), 3), pkg=sum(x["pkg"] for x in summ))
    vessel_name = re.sub(r"\s+(\S+)$", r" V.\1", vessel)
    staff = dict(name=a.name, position=a.position, tel=a.tel)
    write_summary(os.path.join(out, "SUMMARY.pdf"), vessel_name, port, main_code, [x for x in summ if not x.get("empty")], a.gate)
    write_sum_edi(os.path.join(out, "SUM-EDI.pdf"), vessel, port, summ, T, soc, staff)

    print(f"MF: {', '.join(mf)}\nCNTRS: {', '.join(cntrs) or '- (no SOC check)'}\nVESSEL: {vessel}   PORT: {port}")
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
