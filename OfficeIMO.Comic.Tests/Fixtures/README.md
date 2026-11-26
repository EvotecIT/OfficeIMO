Info-ZIP wrote `harbour.cbz` with three independently encoded PNGs
and a ComicInfo sidecar. The archive stores page10 first, then the
cover and page2. The page canvases are 100×150, 120×160 and 144×180.

7-Zip 26.06 (arm64 macOS) wrote `harbour.cb7` from those same four source
files with `7zz a -t7z -m0=LZMA2 -ms=on -mhc=on -mmt=off -mtm=off harbour.cb7 .`.
It uses one solid LZMA2 folder and reports no additional filters.
