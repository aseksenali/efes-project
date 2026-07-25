import os, re, collections

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))

html = open(os.path.join(ROOT, 'system', 'index.dev.html'), encoding='utf-8').read()
print('len:', len(html))

for pat in [r'data-tab="([^"]+)"', r'data-view="([^"]+)"', r'data-page="([^"]+)"',
            r'data-section="([^"]+)"', r'id="view-([^"]+)"', r'id="page-([^"]+)"']:
    m = re.findall(pat, html)
    if m:
        print(pat, '->', sorted(set(m)))

# nav link texts
nav = re.findall(r'<(?:a|button)[^>]*class="[^"]*nav[^"]*"[^>]*>(.*?)</(?:a|button)>', html, re.S)
clean = []
for n in nav:
    t = re.sub(r'<[^>]+>', ' ', n)
    t = re.sub(r'\s+', ' ', t).strip()
    if t:
        clean.append(t)
print('NAV:', clean[:40])

# i18n keys maybe
keys = re.findall(r'\bI18N\b|\blang\b|sidebar', html)
print(collections.Counter(keys))
