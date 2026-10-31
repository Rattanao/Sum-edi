"""Rule checks on the input files, printed as warnings (the documents are still made):
- MF (Excel, the 1) SHED NO. / 2) country / 3) CARGO MOVEMENT / 4) TAX ID lines) - the fixed rules of Check EDI:
  DG at LAEM CHABANG -> SHED 2826, LAOS at BANGKOK -> 0124, USED ENGINE at BANGKOK -> 0126; a TRANSIT sentence
  naming a country -> that country code; TRANSHIPMENT -> (3-TRANSHIPMENT), TRANSIT (not Laos) -> (7-TRANSIT);
  TAX ID name = NOTIFY name.
- MF against CNTRS (Excel): same SHED, DG in both or neither (per container), same 20'/40' size.
- CNTRS: every DG container (SOC too) discharged at LAEM CHABANG under SHED 2826.
Each warning is (level, reference, message): level "ผิด" = wrong, "ตรวจ" = a person should look. DG outside 2826
is "ผิด" when the job uses SHED 2826, else "ตรวจ" (some terminals keep their DG, e.g. 2835).
"""
import os, re
import xlrd

import sum_mf as S
from split_docs import read_cntrs_rows

DEST = {"LAOS": "LA", "LAO": "LA", "MYANMAR": "MM", "CHINA": "CN", "MALAYSIA": "MY", "PHILIPPINES": "PH",
        "CAMBODIA": "KH", "MARSHALL ISLAND": "MH", "INDIA": "IN"}
DG = re.compile(r"\bCLASS\s*[:=]?\s*\d|\bUN\s*#?\s*[:=]?\s*\d{4}\b|\bIMO\b|HAZARD", re.I)
SKIP = ("VESSEL & VOYAGE", "PORT OF LOADING", "BILL OF LANDING NO.", "1) SHIPPER", "3) NOTIFY PARTY")
NOT_NAME = {"CO", "LTD", "LIMITED", "COMPANY", "THE", "PUBLIC", "CORP", "CORPORATION", "INC", "HEAD", "OFFICE", "BRANCH", "NO"}


def words(s):
    """company name as words: letters and digits only, without CO / LTD / HEAD OFFICE ..."""
    return [w for w in re.sub(r"[^A-Z0-9 ]", " ", s.upper()).split() if w not in NOT_NAME and not w.isdigit()]


def same_name(tax, notify):
    """every word of the TAX ID name is in the NOTIFY name; the TAX ID field is often cut off, so its last word
    may be the start of a NOTIFY word (MINBURI, MI)"""
    t, n = words(tax), set(words(notify))
    if t and t[-1] not in n and any(w.startswith(t[-1]) for w in n): t = t[:-1]
    return set(t) <= n


def read_mf_rules(path):
    """MANIFEST .xls -> {B/L: shed, dest, move, tax, notify, text, cntrs, types, pod}"""
    s = xlrd.open_workbook(path, logfile=open(os.devnull, "w")).sheet_by_index(0)
    bls, cur, pod = {}, None, ""
    for r in range(s.nrows):
        v = [str(x).replace("\t", " ").strip() for x in s.row_values(r)] + [""] * 23
        line = " ".join(x for x in v if x)
        if v[0] == "PORT OF LOADING": pod = v[8]
        if "PORT TOTAL" in line or "GRAND TOTAL" in line: cur = None; continue
        if v[0].startswith("PortOfDischarge") or v[0] in SKIP: continue
        if S.BLRE.match(v[0]):
            cur = v[0]; bls[cur] = dict(pod=pod, col0=[], marks=[], desc=[], f15=[], cntrs=[], types=[])
        if cur is None: continue
        b = bls[cur]
        for k, c in (("col0", 0), ("marks", 5), ("marks", 6), ("desc", 10), ("f15", 15)):
            if v[c]: b[k].append(v[c])
        b["cntrs"] += S.CNRE.findall(v[5] + " " + v[6])
        m = re.match(r"^(\d{2}[A-Z]\d)X\d+", v[10])
        if m: b["types"].append(m.group(1))
    for b in bls.values():
        f = " | ".join(b["f15"])
        b["shed"] = (re.search(r"SHED NO\.\s*(\d+)", f) or [None, ""])[1]
        b["dest"] = (re.search(r"2\)[^|]*\(([A-Z]{0,2})\)", f) or [None, ""])[1]
        b["move"] = (re.search(r"CARGO MOVEMENT\s*\(([^)]*)\)", f) or [None, ""])[1]
        b["tax"] = (re.search(r"TAX ID:\s*\((.*)\)", " ".join(b["f15"])) or [None, ""])[1].strip()   # may wrap
        b["text"] = " ".join(b["col0"][1:] + b["marks"] + b["desc"]).upper()
        notify, on = [], False
        for l in b["col0"]:
            if l.startswith("N :"): on, l = True, l[3:]
            elif re.match(r"^[SC] :", l): on = False
            if on: notify.append(l)
        b["notify"] = " ".join(notify)
    return bls


def run_checks(mf, cn, port):
    """-> [(level, reference, message)], [notes about files that could not be checked]"""
    out, notes = [], []
    add = lambda level, ref, msg: out.append((level, ref, msg))
    xmf = [f for f in mf if re.search(r"\.xlsx?$", f, re.I)]
    xcn = [f for f in cn if re.search(r"\.xlsx?$", f, re.I)]
    if len(xmf) < len(mf): notes.append("MF ที่เป็น PDF ไม่ได้ตรวจกฎ (ตรวจได้เฉพาะ MF Excel)")
    if len(xcn) < len(cn): notes.append("CNTRS ที่เป็น PDF ไม่ได้ตรวจกฎ (ตรวจได้เฉพาะ CNTRS Excel)")
    bls = {}
    for f in xmf: bls.update(read_mf_rules(f))
    rows = [r for f in xcn for r in read_cntrs_rows(f)[1]]
    cnt = {r["cntr"]: r for r in rows}
    dg_level = lambda sheds: "ผิด" if "2826" in sheds else "ตรวจ"
    mf_sheds = {b["shed"] for b in bls.values()}
    dg_by_cntr = {}
    for no, b in bls.items():
        t, pod = b["text"], b["pod"] or port
        dg = bool(DG.search(t))
        country = next((k for k in DEST if re.search(rf"TRANS(IT|HIP).*\b{k}\b|\b{k}\b.*TRANS(IT|HIP)", t)), "")
        tranship = "TRANSHIP" in t
        laos = DEST.get(country) == "LA"
        if dg and pod == "LAEM CHABANG" and b["shed"] != "2826":
            add(dg_level(mf_sheds), no, f"สินค้า DG ลง LAEM CHABANG ต้องเป็น SHED NO. 2826 แต่ MF เป็น {b['shed'] or '-'}")
        if laos and pod == "BANGKOK" and b["shed"] != "0124":
            add("ผิด", no, f"LAOS ลง BANGKOK ต้องเป็น SHED NO. 0124 แต่ MF เป็น {b['shed'] or '-'}")
        if "USED ENGINE" in t and pod == "BANGKOK" and b["shed"] != "0126":
            add("ผิด", no, f"USED ENGINE ลง BANGKOK ต้องเป็น SHED NO. 0126 แต่ MF เป็น {b['shed'] or '-'}")
        if country and b["dest"] != DEST[country]:
            add("ผิด", no, f"B/L บอก TRANSIT ไป {country} ประเทศปลายทางต้องเป็น ({DEST[country]}) แต่ MF เป็น ({b['dest']})")
        if tranship and not b["move"].startswith("3"):
            add("ผิด", no, f"มีคำว่า TRANSHIPMENT ต้องเป็น CARGO MOVEMENT (3-TRANSHIPMENT) แต่ MF เป็น ({b['move']})")
        elif country and not laos and not tranship and not b["move"].startswith("7"):
            add("ผิด", no, f"TRANSIT ไป {country} ต้องเป็น CARGO MOVEMENT (7-TRANSIT) แต่ MF เป็น ({b['move']})")
        if b["move"][:1] in ("3", "7", "L") and not country and not tranship and "TRANSIT" not in t:
            add("ตรวจ", no, f"CARGO MOVEMENT ({b['move']}) แต่ไม่พบข้อความ TRANSIT / TRANSHIPMENT ใน B/L")
        if b["tax"] and words(b["tax"]) and not same_name(b["tax"], b["notify"]):
            add("ตรวจ", no, f"TAX ID ({b['tax']}) ไม่ตรงกับ NOTIFY ({b['notify']})")
        for c in b["cntrs"]:                                   # MF against CNTRS, container by container
            x = cnt.get(c)
            if not x: continue
            dg_by_cntr.setdefault(c, []).append((no, dg))
            if x["shed"] and b["shed"] and x["shed"] != b["shed"].zfill(4):
                add("ผิด", f"{no} / {c}", f"SHED NO. ใน MF {b['shed']} ไม่ตรงกับ SHED NUMBER ใน CNTRS {x['shed']}")
            if len(set(b["types"])) == 1:
                csz = "20" if "(20')" in x["size"] else "40" if "(40')" in x["size"] else ""
                if csz and (b["types"][0][0] == "2") != (csz == "20"):
                    add("ผิด", f"{no} / {c}", f"ขนาดตู้ใน MF {b['types'][0]} ไม่ตรงกับ CNTRS {x['size']}")
    for c, lst in dg_by_cntr.items():                          # DG per container: any of its B/L against CNTRS
        cdg, mdg = bool(DG.search(cnt[c]["remark"])), any(d for _, d in lst)
        if cdg != mdg:
            add("ตรวจ", f"{c} ({', '.join(n for n, _ in lst)})", "CNTRS ระบุเป็น HAZARDOUS แต่ใน MF ไม่พบ CLASS / UN" if cdg
                else "MF มี CLASS / UN แต่ CNTRS ไม่ได้ระบุ HAZARDOUS")
    if S.PORTS.get(port) == "THLCH":                            # every DG container in CNTRS, SOC too
        cn_sheds = {r["shed"] for r in rows}
        for x in rows:
            if DG.search(x["remark"]) and x["shed"] != "2826":
                add(dg_level(cn_sheds), x["cntr"], f"CNTRS: ตู้ DG ({' '.join(x['remark'].split())}) ของ "
                                                   f"{' '.join(x['cons'])[:45]} ต้องอยู่ SHED 2826 แต่อยู่ SHED {x['shed'] or '-'}")
    return out, notes
