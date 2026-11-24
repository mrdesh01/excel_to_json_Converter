"""Excel -> data.json, এবং index.html-এর embedded data refresh করে।
ব্যবহার: python convert_excel_to_json.py upload_your_file.xlsx"""
import sys, json, re, datetime
from openpyxl import load_workbook
src = sys.argv[1] if len(sys.argv) > 1 else "upload_your_file.xlsx"
wb = load_workbook(src, data_only=True); ws = wb.active
hdr = [c for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
cats = [(i, h) for i, h in enumerate(hdr) if h]          # category name sits on first col of each 5-col block
num = lambda v: float(v) if isinstance(v, (int, float)) else None
recs = []
for r in ws.iter_rows(min_row=3, values_only=True):
    if not r[0] or not r[1]: continue
    c = {}
    for i, name in cats:
        c[name.title() if False else name] = {"tgt": num(r[i]) or 0, "ims": num(r[i+1]), "do": num(r[i+2]) or 0, "del": num(r[i+3]) or 0}
    recs.append({"region": r[0], "area": r[1], "categories": c})
names = [n for _, n in cats]
data = {"meta": {"source": src.split("/")[-1], "sheet": ws.title, "generated_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "categories": names, "beverage_components": ["CSD", "HOTFILL", "HDPE"],
        "top_level": [n for n in names if n not in ("CSD", "HOTFILL", "HDPE")]}, "records": recs}
json.dump(data, open("data.json", "w"), ensure_ascii=False, indent=1)
try:
    h = open("index.html", encoding="utf-8").read()
    h = re.sub(r"/\*DATA_START\*/.*?/\*DATA_END\*/", lambda m: "/*DATA_START*/window.__DATA__=" + json.dumps(data, ensure_ascii=False) + ";/*DATA_END*/", h, flags=re.S)
    open("index.html", "w", encoding="utf-8").write(h)
except FileNotFoundError: pass
print(len(recs), "areas,", len(names), "categories ->", "data.json")
