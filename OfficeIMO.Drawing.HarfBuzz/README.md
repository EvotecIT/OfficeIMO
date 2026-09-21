# OfficeIMO.Drawing.HarfBuzz

`OfficeIMO.Drawing.HarfBuzz` is the optional full OpenType shaping adapter for
OfficeIMO renderers. It implements the shared
`IOfficeTextShapingProvider` contract with HarfBuzz GSUB/GPOS processing while
keeping `OfficeIMO.Core` and `OfficeIMO.Pdf` dependency-light. Drawing APIs remain
in the `OfficeIMO.Drawing` namespace.

```csharp
using OfficeIMO.Drawing.HarfBuzz;
using OfficeIMO.Pdf;

PdfOptions options = new PdfOptions()
    .SetTextShapingProvider(OfficeHarfBuzzTextShapingProvider.Instance);
```

Use this package when documents contain scripts, combining marks, contextual
forms, kerning, or font substitutions that need a complete OpenType shaping
engine. The core packages retain their bounded managed fallback when this
adapter is not installed.

The adapter also carries HarfBuzz top-to-bottom advances through the shared
glyph contract. `OfficeDrawing.AddVerticalText(...)` uses those advances for
raster outlines when this package is registered. SVG keeps the searchable
logical string and declares browser-native vertical shaping. PDF paints
positioned embedded glyphs when the provider supplies complete vertical
advances and logical coverage and the writer can align their ink. Otherwise
it paints searchable stacked text, reports an approximation, and strict
conversion profiles reject it. Without the adapter, the managed provider
declines true vertical shaping and the same fallback is reported.

The package uses HarfBuzzSharp `14.2.1.2` and matching Windows, Linux, macOS,
and WebAssembly native assets. HarfBuzzSharp is MIT licensed; applications
should include the upstream notices required by their own distribution policy.
