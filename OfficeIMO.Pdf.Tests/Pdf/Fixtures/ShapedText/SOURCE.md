# Painted-glyph rendering references

`PdfPaintedGlyphRenderingTests` renders page 1 of each case at 144 dpi and requires at least
0.990 ink precision and recall (1 px tolerance) against the Poppler raster stored here. Poppler
is an independent renderer; a PDF paints exact shaped glyphs, so both renders must agree.
`PdfArabicPaintedFormsSceneTests` checks that the page scene of `cairo-arabic-extended.pdf` names
each painted letter by the presentation form the Unicode joining rules select.

| Reference | Source PDF | Contract |
| --- | --- | --- |
| `poppler-chrome-arabic.png` | `chrome-arabic.pdf` | Chrome/Skia output: visual-order single-glyph runs, ToUnicode presentation forms and base letters, and font-specific contextual alternates (lam before ra). |
| `poppler-chrome-arabic-extended.png` | `chrome-arabic-extended.pdf` | Chrome/Skia output with Persian and Urdu letters, line-initial glyphs whose ink reaches the page clip, and one word painted across two font subsets. |
| `poppler-cairo-arabic-extended.png` | `cairo-arabic-extended.pdf` | Pango/Cairo output with Persian and Urdu letters and one word whose middle letters use DejaVu Sans; ToUnicode maps every contextual glyph to its base letter. |
| `poppler-chrome-devanagari.png` | `chrome-devanagari.pdf` | Chrome/Skia output with multi-glyph Devanagari runs, conjunct glyphs mapped to clusters, and glyphs whose ToUnicode text is U+0000 (the text is carried by ActualText). |
| `poppler-space-coded-glyph.png` | `space-coded-glyph.pdf` | An inked glyph whose ToUnicode text is a space, painted over a letter like a colour-font layer, next to a real space glyph and at both ends of multi-glyph text runs. |
| `poppler-ligature-run.png` | `ligature-run.pdf` | Multi-glyph Latin runs containing ligature glyphs mapped to their letters. |
| `poppler-ghostscript-ligature-run.png` | `ghostscript-ligature-run.pdf` | Ghostscript redistillation of `ligature-run.pdf`: its ToUnicode maps the ffi ligature to `f` plus U+00CF, so the glyph must be drawn by id. |
| `poppler-cairo-rtl-0.png` | `OfficeIMO.TestAssets/MultilingualLayout/rtl-0-native.pdf` | Pango/Cairo output: CIDFontType2 subsets without cmap, logical-order TJ glyph runs whose ToUnicode maps every contextual form to its base letter. |
| `poppler-cairo-rtl-90.png` | `OfficeIMO.TestAssets/MultilingualLayout/rtl-90-native.pdf` | The same content on a page rotated 90 degrees. |
| `poppler-cairo-latin-0.png` | `OfficeIMO.TestAssets/MultilingualLayout/latin-0-native.pdf` | CIDFontType2 subsets without cmap, including an `ffi` ligature glyph. |
| `poppler-word-mac-report.png` | `ReferenceBaselines/microsoft-word-16.109-native-word-report.pdf` | Simple TrueType subsets with only a (1,0) cmap and TJ runs with character spacing. |
| `poppler-word-windows-summary.png` | `ReferenceBaselines/microsoft-word-windows-word-business-delivery-summary.pdf` | Word word-final spaces painted inside TJ strings. |

| Source PDF | SHA-256 | Producer |
| --- | --- | --- |
| `chrome-arabic.pdf` | `406b12a79ffbf5d9957b081c6a368b4dc26ea2bbc2c74da6aa8ef29d737cddb5` | Chrome 153.0.8010.53 headless `--print-to-pdf` of `chrome-arabic.html` |
| `chrome-arabic-extended.pdf` | `962de78f78f8e115ccdea4a78d99f553449ecfe792d6f7bd8ae1053af7460ed3` | Chrome 153.0.8010.53 headless `--print-to-pdf` of `chrome-arabic-extended.html` |
| `chrome-devanagari.pdf` | `11c9a02849658bbba52d08eb5be61aca37028f92477167951ac82b0f2f67ad6a` | Chrome 153.0.8010.53 headless `--print-to-pdf` of `chrome-devanagari.html` |
| `cairo-arabic-extended.pdf` | `96b1b29603c2937c194c000bc05da1db856629f5f675e671ee01fb85f170c38a` | `generate_cairo_extended.py` with Cairo 1.18.0 and Pango 1.52.1 (Ubuntu) |
| `space-coded-glyph.pdf` | `cc42204fe1329b4077c463d6b4258c5a6c13ae14286bf1a573c54bda4f0a4089` | `create_space_coded_glyph_fixture.py` |
| `ligature-run.pdf` | `224fec268284175e96103e12b76e89b2ad7261ef120435d8b6cb50d25c65f259` | `create_ligature_run_fixture.py` |
| `ghostscript-ligature-run.pdf` | `07c80e9df7b07cbe997aaf4d22886c721bbed93244d0ed7194cdfbf70a3d9b68` | GPL Ghostscript 10.02.1 (Ubuntu) `pdfwrite` of `ligature-run.pdf` |

Fonts: the Arabic fixtures embed subsets of the repository's
`Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/NotoSansArabic-Regular.ttf` (SIL Open Font
License, see `OFL-Noto.txt` beside it); `chrome-arabic-extended.pdf` also embeds a subset of a
renamed copy of that font. `cairo-arabic-extended.pdf` embeds a subset of DejaVu Sans 2.37
(Bitstream Vera and public-domain DejaVu changes; redistribution permitted). `chrome-devanagari.pdf`
embeds a subset of `OfficeIMO.Drawing.Tests/TestAssets/NotoSansDevanagari-Regular.ttf` (SIL Open Font
License, see `OFL-1.1.txt` beside it). `space-coded-glyph.pdf` uses OfficeIMO Baseline Sans and the
ligature fixtures use Carlito (SIL Open Font License, see the `OFL-Carlito.txt` files).

Regenerate from the repository root with Poppler `pdftoppm` and Python Pillow. `--chrome`
reprints the Chrome PDFs first (it also needs fontTools); Chrome output is not byte-reproducible,
so update the hashes above when reprinting:

```sh
python3 OfficeIMO.Pdf.Tests/Pdf/Fixtures/ShapedText/generate_cairo_extended.py
python OfficeIMO.Pdf.Tests/Pdf/Fixtures/ShapedText/create_space_coded_glyph_fixture.py
python OfficeIMO.Pdf.Tests/Pdf/Fixtures/ShapedText/create_ligature_run_fixture.py
gs -q -dNOPAUSE -dBATCH -dSAFER -sDEVICE=pdfwrite -sOutputFile=OfficeIMO.Pdf.Tests/Pdf/Fixtures/ShapedText/ghostscript-ligature-run.pdf OfficeIMO.Pdf.Tests/Pdf/Fixtures/ShapedText/ligature-run.pdf
python OfficeIMO.Pdf.Tests/Pdf/Fixtures/ShapedText/create_references.py --chrome "<path to chrome>"
```
