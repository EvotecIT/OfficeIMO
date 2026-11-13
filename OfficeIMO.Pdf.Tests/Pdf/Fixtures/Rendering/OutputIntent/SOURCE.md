# Output-intent composition corpus

`generate.py` creates four 128 x 96 point pages with ReportLab 5.1.0 and an embedded Carlito font from this repository. The pages exercise opaque paint, fill opacity, Multiply blending, and a PNG image soft mask. `reportlab-authored.pdf` is the independent producer output. pypdf 7.0.1 adds only a catalog output intent to create `reportlab-output-intent.pdf`; neither PDF is produced by OfficeIMO.

LittleCMS creates a narrow-gamut RGB matrix/TRC destination profile. Its fixed creation timestamp makes regeneration repeatable. The narrow gamut makes proofing before and after overlapping paint observably different. This is a rendering fixture, not a PDF/A conformance claim.

MuPDF 1.28.5 renders the authored PDF at 72 dpi over white. Pillow 13.0.0 / LittleCMS 2.19 then converts the composed sRGB pixels to bounded 8-bit destination RGB and back to sRGB with relative colorimetry. The explicit device stage retains gamut clipping that a fused floating-point matrix transform can eliminate. The PNGs and `reference.json` are independent reference output; product tests never invoke these tools. Interior samples allow four byte levels for producer rounding and the intermediate device quantization. Text labels, their (6, 82) point baselines, drawing cloning, nested composition, SVG image dimensions, page diagnostics, and bounded display tiles have separate assertions.

Regenerate in an isolated Python environment with `reportlab==5.1.0`, `pypdf==7.0.1`, and `Pillow==13.0.0`, with LittleCMS and `mutool` available:

```sh
python generate.py
```

Generated fixture content is covered by the repository license. Carlito is embedded under its existing SIL Open Font License; see the font assets and their notices in `Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts`.

| Artifact | SHA-256 |
| --- | --- |
| `page-2-reference.png` | `d26b4a9b9211ba0cc6d292ed5c71f094ec61b39a4b0526d592fb213799170580` |
| `page-3-reference.png` | `43b4df7da4145ca9d840330abefcf451cf793bc57283cecca618d2dbeafb0604` |
| `page-4-reference.png` | `3a067d2441329ca21a8bf855b189e802e3a84fe44bcdc5498858ba280de19098` |
| `reference.json` | `27bd5056a5e7a2362364597b09c7262c3195a0a55a647b431aca863afdf3831c` |
| `reportlab-authored.pdf` | `edcb9a07cbad4a64ee7789c150845413f9b39c602e8131e04b921a8336963729` |
| `reportlab-output-intent.pdf` | `e8998154e70fc24a5f4a58d33a11e04bb394a142b566467586ed01c91e3ea292` |
