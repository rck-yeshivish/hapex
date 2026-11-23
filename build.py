#!/usr/bin/env python3
"""Usage: python build.py [שרשים חד פעמים.xlsx] [index.html]
Reads the spreadsheet (keeping bold formatting in the verse column) and
rewrites the data block inside index.html between the DATA markers."""
import sys, re, json
import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from books import parse_ref, ORDER, SECTION, BOOKS

XLSX = sys.argv[1] if len(sys.argv) > 1 else "שרשים חד פעמים.xlsx"
HTML = sys.argv[2] if len(sys.argv) > 2 else "index.html"

# Cells where the sheet lacks bold formatting or a citation; filled in by hand.
OVERRIDE_WORD = {"בלס": "וּבוֹלֵס", "טאב": "טְאֵב", "ית": "יָתְהוֹן", "סמן": "נִסְמָן", "רטפש": "רֻטְפַשׁ"}
OVERRIDE_REF = {"חיט": "עזרא ד:יב"}

NIQ = re.compile(r"[\u0591-\u05C7]")
def plain(s): return NIQ.sub("", s)
def clean(s):
    return None if s is None or str(s).strip() == "" else str(s).strip()

def verse_parts(cell_value, cell_bold, root):
    """Return list of [text, isBold] segments."""
    if isinstance(cell_value, CellRichText):
        parts = []
        for b in cell_value:
            if isinstance(b, TextBlock):
                parts.append([b.text, bool(b.font.b)])
            else:  # plain run inherits the cell font
                parts.append([str(b), bool(cell_bold)])
    else:
        parts = [[str(cell_value), False]]
    return [p for p in parts if p[0]]

def tidy_parts(parts):
    # merge neighbours with same flag, trim edges
    out = []
    for t, b in parts:
        if out and out[-1][1] == b: out[-1][0] += t
        else: out.append([t, b])
    if out: out[0][0] = out[0][0].lstrip(); out[-1][0] = out[-1][0].rstrip()
    return out

def rishon(val, root, kind):
    """kind: 'root' (E-H), 'phrase' (Rashi / Ibn Ezra)"""
    v = clean(val)
    if v is None: return None
    note = "~" in v
    star = "*" in v
    core = v.replace("~", "").replace("*", "").strip()
    if kind == "root":
        if core == root and star: st = "star"        # same root, but others exist
        elif core == root:        st = "hapax"
        else:                     st = "other"       # different root
    else:
        st = "other" if core == "צרב" else "hapax"   # phrase = says unique
    return {"v": core, "s": st, "c": note}

wb = openpyxl.load_workbook(XLSX, rich_text=True)
ws = wb.active
entries = []
problems = []
for n, row in enumerate(ws.iter_rows(min_row=2), start=2):
    vals = [c.value for c in row]
    if not clean(vals[1]): continue
    root = clean(vals[1])
    parts = tidy_parts(verse_parts(vals[2], row[2].font.b, root)) if vals[2] else []
    ref = clean(vals[3]) or OVERRIDE_REF.get(root)
    pr = parse_ref(ref) if ref else None
    if not pr: problems.append((n, root, ref)); continue
    book, ch, vs = pr
    # bold word
    if not any(b for _, b in parts) and root in OVERRIDE_WORD:
        w = OVERRIDE_WORD[root]
        full = "".join(t for t, _ in parts)
        i = full.find(w)
        if i >= 0:
            parts = [[full[:i], False], [w, True], [full[i+len(w):], False]]
            parts = [p for p in parts if p[0]]
        else: problems.append((n, root, "override word not found"))
    word = "".join(t for t, b in parts if b).strip()
    # position of the bold word in the verse
    words = "".join(t for t, _ in parts).split()
    pos = None
    if word and parts:
        before = "".join(t for t, b in parts[:next((i for i, p in enumerate(parts) if p[1]), 0)])
        k = len(before.split()); wl = len(word.split())
        pos = "only" if len(words) == wl else "first" if k == 0 else "last" if k + wl >= len(words) else "middle"
    entries.append({
        "id": n, "root": root, "aram": clean(vals[0]) == "#",
        "parts": parts, "word": word, "ref": ref, "book": book, "ch": ch, "vs": vs,
        "sec": SECTION[book], "ord": ORDER[book] * 1_000_000 + ch * 1000 + vs,
        "pos": pos, "len": len(plain(root)),
        "r": {
            "ryach": rishon(vals[4], root, "root"), "menachem": rishon(vals[5], root, "root"),
            "ribag": rishon(vals[6], root, "root"), "radak": rishon(vals[7], root, "root"),
            "rashi": rishon(vals[8], root, "phrase"), "ibnezra": rishon(vals[9], root, "phrase"),
        },
        "notes": clean(vals[10]) or "",
    })

for e in entries:
    e["agree"] = sum(1 for r in e["r"].values() if r and r["s"] == "hapax")

blob = json.dumps({"entries": entries, "books": [b[1] for b in BOOKS],
                   "sections": {b[1]: b[0] for b in BOOKS}}, ensure_ascii=False, separators=(",", ":"))
blob = blob.replace("</", "<\\/")
html = open(HTML, encoding="utf-8").read()
new = re.sub(r"(<!--DATA-->\s*<script id=\"data\" type=\"application/json\">).*?(</script>\s*<!--/DATA-->)",
             lambda m: m.group(1) + blob + m.group(2), html, flags=re.S)
open(HTML, "w", encoding="utf-8").write(new)
print(f"{len(entries)} roots written to {HTML}")
for p in problems: print("PROBLEM:", p)
print("no bold word:", [e["root"] for e in entries if not e["word"]])
print("positions:", {k: sum(1 for e in entries if e["pos"] == k) for k in ("first","middle","last","only",None)})
