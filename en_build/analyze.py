# -*- coding: utf-8 -*-
"""Show each paragraph's runs with an rPr fingerprint, so we can see which
adjacent runs share formatting (mergeable) vs. differ (intentionally styled)."""
import glob, os, re, hashlib
from xml.dom import minidom

SLIDES = r"C:\Users\Aset\Downloads\Telegram Desktop\efes (5)\efes\en_build\unpacked\ppt\slides"
out = []

def local(n): return n.localName or n.tagName

def child(parent, name):
    for c in parent.childNodes:
        if c.nodeType == c.ELEMENT_NODE and local(c) == name:
            return c
    return None

def runs_of(p):
    return [c for c in p.childNodes if c.nodeType == c.ELEMENT_NODE and local(c) == "r"]

def text_of(r):
    t = child(r, "t")
    if t is None: return None
    return t.firstChild.data if t.firstChild else ""

def rpr_fp(r):
    rpr = child(r, "rPr")
    if rpr is None: return "none"
    x = rpr.toxml()
    return hashlib.md5(x.encode("utf-8")).hexdigest()[:6]

files = sorted(glob.glob(os.path.join(SLIDES, "slide*.xml")),
               key=lambda p: int(re.search(r"slide(\d+)\.xml", p).group(1)))

for f in files:
    name = os.path.basename(f)
    dom = minidom.parseString(open(f, encoding="utf-8").read())
    paras = [p for p in dom.getElementsByTagName("a:p")]
    out.append("===== %s : %d paragraphs =====" % (name, len(paras)))
    for pi, p in enumerate(paras):
        rs = runs_of(p)
        rs = [r for r in rs if text_of(r) is not None]
        if not rs:
            continue
        if len(rs) == 1:
            out.append("  P%02d [1 run  fp=%s] %r" % (pi, rpr_fp(rs[0]), text_of(rs[0])))
        else:
            fps = [rpr_fp(r) for r in rs]
            uniform = len(set(fps)) == 1
            out.append("  P%02d [%d runs %s]" % (pi, len(rs), "UNIFORM" if uniform else "MIXED"))
            for r in rs:
                out.append("        fp=%s %r" % (rpr_fp(r), text_of(r)))
    out.append("")

dst = r"C:\Users\Aset\Downloads\Telegram Desktop\efes (5)\efes\en_build\paras.txt"
open(dst, "w", encoding="utf-8").write("\n".join(out))
print("WROTE paras.txt")
