# Vertical PDF drawing references

The opt-in `validate_references.py` gate checks native PDF drawing against two
independent CJK PDF producers. From the repository root, run
`python OfficeIMO.Pdf.Tests/Pdf/Fixtures/Vertical/validate_references.py`.
It needs .NET 10, Poppler's
`pdffonts`, `pdftotext`, and `pdftoppm`, and Python Pillow. Temporary files are
removed after the check; pass `--output <directory>` to retain them.

| Reference | Source and provenance | SHA-256 | Contract |
| --- | --- | --- | --- |
| `chrome-identity-h.pdf` | Chrome/Skia PDF m153 print of the adjacent `chrome-identity-h.html`, using the repository's `NotoSansJP-OfficeIMO-Common.ttf` | `c10e749fa6601a6c4074e20a4b43fd6c4d993ae7c87d52a0a5002e4ef8c7d182` | Embedded `/Identity-H`, independent glyph positioning, ToUnicode, Japanese vertical punctuation and 120 × 180 pt page. |
| `vertical.pdf` | Mozilla [pdf.js test corpus](https://github.com/mozilla/pdf.js/blob/d54c193bd4dd6c34759cb88f1a3b78db66f0962c/test/pdfs/vertical.pdf), commit `d54c193bd4dd6c34759cb88f1a3b78db66f0962c`, produced by dvipdfmx 20090506 | `514511143db12309893fb69cdf98c76f2361c20da394f79fdff2c119ed7a4393` | Embedded `/Identity-V` CJK font; page 1 has two top-to-bottom Japanese columns, ordered right to left. Downloaded only by the opt-in gate, with an exact 6,905-byte limit. |

The gate requires Poppler to extract the exact logical text from both references
and the OfficeIMO output. It checks the `/Identity-V` reference's column
geometry from word boxes and compares the OfficeIMO and Chrome rendered ink at
72 dpi. Their full-page dark-pixel overlap must be at least 0.90; the pinned
reference run produced 0.936. The normal HarfBuzz integration test separately
checks vertical substitutions, each emitted glyph ID and text matrix against
the provider's X/Y offsets and advances, `/ActualText` readback, and
`RequireNoLoss()`. Separate tests require a diagnosed fallback for partial
shaping results and drawing contexts without a logical-text wrapper.

The dvipdfmx file has no ToUnicode map. Poppler recovers its Japanese text from
the predefined CJK font mapping; OfficeIMO's PDF reader currently does not.
This reference therefore qualifies the drawing writer's vertical geometry and
independent extraction evidence, not import or structured-read support for
unmapped `/Identity-V` fonts. The downloaded fixture is not redistributed
because its embedded font's redistribution terms have not been established.
