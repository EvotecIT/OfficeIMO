PNG color declarations and native defaults

These nine 16x16 PNGs are generated test inputs with no third-party image assets. `generate_reference.py` writes `gAMA` and optional `cHRM` chunks for RGB, grayscale, indexed color and alpha. The native expected columns in `expected.csv` retain the source channels because ECMA-388 15.3.7 M8.30 and Table 15-3 require sRGB defaults for integer samples without a usable ICC profile. A producer must associate/embed a usable ICC profile when those defaults do not describe its intended colors (M8.43).

Eight cases have separate calibrated columns that record Pillow 12.3.0 / LittleCMS 2.19 conversion through a constructed RGB power-curve profile with relative colorimetric intent. They demonstrate that general PNG calibration differs from XPS native defaults; they are not the native renderer's expected colors. Alpha is unchanged by either color transform. The ninth image declares canonical sRGB cICP alongside linear gamma; PNG 3 cICP precedence and the native defaults both retain its channel values. Its calibrated comparison columns are empty because no synthesized ICC reference is used. The cases exercise linear/power gamma, sRGB, Display P3 and Adobe-style primaries, D50 white, grayscale and indexed transparency.

The optional generator uses Pillow and its bundled LittleCMS library outside production code and normal build/test requirements. The `.dylibs` lookup describes the macOS reference environment; use `--lcms-library` for an equivalent test-only runtime when regenerating elsewhere. Checked-in inputs and expectations need no native color engine during tests.

References: [ECMA-388](https://ecma-international.org/publications-and-standards/standards/ecma-388/), [PNG color declarations](https://www.w3.org/TR/png-3/#11gAMA), [LittleCMS](https://github.com/mm2/Little-CMS).

`apng-static.png` and `apng-static-embedded.png` are generated one-pixel PNG
inputs with opaque red static IDAT data and a separate opaque green APNG frame.
They use standard PNG CRC-32 and zlib scanline encoding. The second embeds the
existing `IccColorCorpus/littlecms-rgb-matrix.icc` test profile, whose provenance
and license live in that corpus. The tests also add a gamma chunk or associate
that profile to cover all native preparation routes. These inputs protect
static PNG paint selection; they do not qualify APNG playback.
