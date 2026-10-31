"""ATTACHED sheets: any other file in the input folder (PDF, Excel, picture, any name) that is not the MF / CNTRS.
Its B/L comes from the content ("BL NO. HASL..." with "VESSEL ...", e.g. *AS PER ATTACHED SHEET*), page by page for a
PDF (a page without BL NO. continues the page before), else from a HASL... number in the file name. Each page is
put into MF-<DELIVERY>.pdf right after the MF page where its B/L ends; only pages whose B/L is in the MF and whose
VESSEL (when written) is the MF vessel. Problems are reported by checks.py.
"""
import os, re, glob
import xlrd, pymupdf

import sum_mf as S

RE_BLNO = re.compile(r"\bB\s*/?\s*L\s*NO\.?\s*:?\s*(HASL[A-Z0-9]{10,14})")
RE_VESSEL = re.compile(r"\bVESSEL\b(?:\s*(?:NAME|&\s*VOYAGE))?\s*:?\s*([^\n]+)")
RE_HASL = re.compile(r"HASL[A-Z0-9]{10,14}", re.I)
PICTURE = (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff")


def same_vessel(a, b):
    """SAWASDEE VEGA 2613S = SAWASDEE VEGA V.2613S = SAWASDEE  VEGA 2613S"""
    n = lambda s: re.sub(r"[^A-Z0-9]", "", re.sub(r"\bV\.?\s*(?=\d)", "", s.upper()))
    return n(a) == n(b)


def name_bl(path):
    m = RE_HASL.search(os.path.basename(path))
    return m.group(0).upper() if m else ""


def sheet_lines(path):
    """Excel -> text lines (the cells of a row joined), every sheet"""
    out = []
    if path.lower().endswith(".xls"):
        wb = xlrd.open_workbook(path, logfile=open(os.devnull, "w"))
        for s in wb.sheets():
            for r in range(s.nrows): out.append("   ".join(str(x).strip() for x in s.row_values(r) if str(x).strip()))
    else:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                out.append("   ".join(str(x).strip() for x in row if x is not None and str(x).strip()))
    while out and not out[-1]: out.pop()
    return out


def text_pages(lines):
    """text lines -> a PDF of A4 pages (Times 9), for an Excel ATTACHED sheet"""
    doc = pymupdf.open()
    for i in range(0, max(len(lines), 1), 70):
        pg = doc.new_page(width=595, height=842)
        for k, l in enumerate(lines[i:i + 70]):
            if l: pg.insert_text((40, 50 + k * 11), l[:150], fontname="tr", fontfile=r"C:\Windows\Fonts\times.ttf", fontsize=9)
    return doc


def read_one(path):
    """one file -> pymupdf doc of its pages, [dict(page, vessel, bl, src)] ; doc None = not usable"""
    low = path.lower()
    try:
        if low.endswith(".pdf"):
            doc = pymupdf.open(path); texts = [pg.get_text() for pg in doc]
        elif re.search(r"\.xlsx?$", low):
            lines = sheet_lines(path); doc = text_pages(lines); texts = ["\n".join(lines)] + [""] * (len(doc) - 1)
        elif low.endswith(PICTURE):
            img = pymupdf.open(path); rect = img[0].rect
            doc = pymupdf.open(); pg = doc.new_page(width=595, height=842)
            pg.insert_image(pymupdf.Rect(20, 20, 575, 822), filename=path, keep_proportion=True); texts = [""]
        else:
            return None, []
    except Exception:
        return None, []
    pages, vessel, bl = [], "", ""
    nb = name_bl(path)
    for i, t in enumerate(texts):
        m = RE_BLNO.search(t)
        if m: bl = m.group(1)
        m = RE_VESSEL.search(t)
        if m and m.group(1).strip(): vessel = m.group(1).strip()
        pages.append(dict(page=i, vessel=vessel, bl=bl or nb, src="content" if bl else "name" if nb else ""))
    return doc, pages


def plan(folder, used, vessel, bls):
    """every other file in the folder -> {B/L: [(doc, page)]} to put into the MF, [all pages read (for checks)]"""
    use, pages = {}, []
    skip = {os.path.abspath(f) for f in used}
    for f in sorted(glob.glob(os.path.join(folder, "*"))):
        if not os.path.isfile(f) or os.path.abspath(f) in skip or os.path.basename(f).startswith(("~$", ".")): continue
        if re.search(r"\.(xlsx?|pdf)$", f, re.I) and S.kind_of(f)[0] in ("mf", "cntrs"): continue   # other voyages
        doc, ps = read_one(f)
        if doc is None:
            pages.append(dict(file=f, page=-1, vessel="", bl="", src="")); continue
        for p in ps:
            p["file"] = f; pages.append(p)
            if p["bl"] in bls and (not p["vessel"] or same_vessel(p["vessel"], vessel)):
                use.setdefault(p["bl"], []).append((doc, p["page"]))
    return use, pages
