Info-ZIP wrote `harbour.cbz` with three independently encoded PNGs
and a ComicInfo sidecar. The archive stores page10 first, then the
cover and page2. The page canvases are 100×150, 120×160 and 144×180.

7-Zip 26.06 (arm64 macOS) wrote `harbour.cb7` from those same four source
files with `7zz a -t7z -m0=LZMA2 -ms=on -mhc=on -mmt=off -mtm=off harbour.cb7 .`.
It uses one solid LZMA2 folder and reports no additional filters.

RARLAB RAR 7.23 (arm64 macOS) wrote `harbour.cbr` with
`rar a -ma5 -m0 harbour.cbr ComicInfo.xml cover.png page10.png page2.png`.
All members are stored at the root. The producer's `rar t harbour.cbr` passed.

[RARLAB's fixture manifest](rar5-provenance.json) records the stored and
compressed CBR commands, producer version and SHA-256 checksums. To regenerate
the RAR samples, run `Build/Generate-Rar5Fixtures.py` with paths to the official
RAR executables.
