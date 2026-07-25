# -*- coding: utf-8 -*-
import re, glob, os, json

SLIDES = r"C:\Users\Aset\Downloads\Telegram Desktop\efes (5)\efes\en_build\unpacked\ppt\slides"
out_lines = []
all_unique = {}

files = sorted(glob.glob(os.path.join(SLIDES, "slide*.xml")),
               key=lambda p: int(re.search(r"slide(\d+)\.xml", p).group(1)))

pat = re.compile(r"<a:t>(.*?)</a:t>", re.DOTALL)
for f in files:
    name = os.path.basename(f)
    with open(f, encoding="utf-8") as fh:
        xml = fh.read()
    texts = pat.findall(xml)
    out_lines.append("===== %s (%d text nodes) =====" % (name, len(texts)))
    for i, t in enumerate(texts):
        out_lines.append("[%02d] %r" % (i, t))
        all_unique.setdefault(t, []).append(name)
    out_lines.append("")

out_lines.append("===== UNIQUE STRINGS (%d) =====" % len(all_unique))
for t, where in all_unique.items():
    out_lines.append("%r   <- %s" % (t, ",".join(sorted(set(where)))))

dst = r"C:\Users\Aset\Downloads\Telegram Desktop\efes (5)\efes\en_build\strings.txt"
with open(dst, "w", encoding="utf-8") as fh:
    fh.write("\n".join(out_lines))

# also dump unique list as JSON for building the map
with open(r"C:\Users\Aset\Downloads\Telegram Desktop\efes (5)\efes\en_build\unique.json", "w", encoding="utf-8") as fh:
    json.dump(list(all_unique.keys()), fh, ensure_ascii=False, indent=1)

print("WROTE strings.txt and unique.json; unique count:", len(all_unique))
