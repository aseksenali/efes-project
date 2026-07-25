# RU -> EN deck translation

The English base deck is produced from the Russian one by replacing every
`<a:t>` text node with an exact-match translation.

1. Unpack the Russian base deck:

       cd deck/build && mkdir -p gen/unpacked
       (cd gen/unpacked && unzip -o ../../../base/Efes_Nexus_RU_base.pptx)

2. `python3 extract.py` — dumps every text node to `strings.txt` and the unique
   list to `unique.json` (both are scratch output, not tracked).

3. `python3 apply.py` — rewrites the unpacked slide XML in place using the
   per-slide translation maps. **`apply.py` is the source of truth for the
   translations**; it reports any text node the map does not cover, so coverage
   should always come back at 100%.

4. Re-zip `gen/unpacked` into `deck/base/Efes_Nexus_EN_base.pptx`.

Only needed when `build.js` changes the slide text. Day-to-day edits go through
`deck/build/transform.py`, which already handles both languages.
