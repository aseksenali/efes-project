# -*- coding: utf-8 -*-
"""Translate Russian document-property metadata in core.xml and app.xml."""
import os, sys
sys.stdout.reconfigure(encoding="utf-8")
BASE = r"C:\Users\Aset\Downloads\Telegram Desktop\efes (5)\efes\en_build\unpacked\docProps"

def patch(path, repls):
    p = os.path.join(BASE, path)
    s = open(p, encoding="utf-8").read()
    for a, b in repls:
        n = s.count(a)
        s = s.replace(a, b)
        print("%s: %r -> %r  (%d)" % (path, a, b, n))
    open(p, "w", encoding="utf-8").write(s)

patch("core.xml", [
    ("Efes Nexus — бизнес-кейс", "Efes Nexus — business case"),
])
patch("app.xml", [
    ("Широкоэкранный", "Widescreen"),
    ("Использованные шрифты", "Fonts Used"),
    ("Заголовки слайдов", "Slide Titles"),
    ("Тема", "Theme"),
    ("Презентация PowerPoint", "PowerPoint Presentation"),
])
print("done")
