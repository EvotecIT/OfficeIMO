# Output-intent composition corpus

`generate.py` creates five 128 x 96 point pages with ReportLab 5.1.0 and an embedded Carlito font from this repository. The pages exercise opaque paint, fill opacity, Multiply blending, a PNG image soft mask, and an otherwise opaque page declared as a transparency group. `reportlab-authored.pdf` is the independent producer output. pypdf 7.0.1 adds a catalog output intent and a transparency-group declaration on the final solid-paint page to create `reportlab-output-intent.pdf`; neither PDF is produced by OfficeIMO.

LittleCMS creates a narrow-gamut RGB matrix/TRC destination profile. Its fixed creation timestamp makes regeneration repeatable. The narrow gamut makes proofing before and after overlapping paint observably different. This is a rendering fixture, not a PDF/A conformance claim.

MuPDF 1.28.5 renders the authored PDF at 72 dpi over white. Pillow 13.0.0 / LittleCMS 2.19 then converts the composed sRGB pixels to bounded 8-bit destination RGB and back to sRGB with relative colorimetry. The explicit device stage retains gamut clipping that a fused floating-point matrix transform can eliminate. The PNGs and `reference.json` are independent reference output; product tests never invoke these tools. Reference PNGs retain only sRGB pixel data, without timestamp-bearing ICC metadata. Interior samples allow four byte levels for producer rounding and the intermediate device quantization. Text labels, their (6, 82) point baselines, drawing cloning, nested composition, SVG image dimensions, page diagnostics, and bounded display tiles have separate assertions. The final solid-paint page verifies that portable JSON rejects the unrepresentable color transform rather than silently dropping it.

Regenerate in an isolated Python environment with `reportlab==5.1.0`, `pypdf==7.0.1`, and `Pillow==13.0.0`, with LittleCMS and `mutool` available:

```sh
python generate.py
```

Generated fixture content is covered by the repository license. Carlito is embedded under its existing SIL Open Font License; see the font assets and their notices in `Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts`.

| Artifact | SHA-256 |
| --- | --- |
| `page-2-reference.png` | `4f5806a5213464f477d10f872013165c1ab4a2d0a3919b64c628e816bdae3e07` |
| `page-3-reference.png` | `8db014033915039deda29e443872e31f8979a366011787827f347c4558b691fb` |
| `page-4-reference.png` | `c146d81b92d03a34526dcfded4cf3929ecd8f0bdeb240600c68993da99d63f25` |
| `reference.json` | `27bd5056a5e7a2362364597b09c7262c3195a0a55a647b431aca863afdf3831c` |
| `reportlab-authored.pdf` | `6ee10c5bd68f749a1f93cb923d668df7eb14fa3b6fc347fc4696ab5d876f01ef` |
| `reportlab-output-intent.pdf` | `021845e1c099d34d636c21ea0d3221b9a97057fb0d8a5d8d32be053038954ea2` |
