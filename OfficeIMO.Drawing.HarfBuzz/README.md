# OfficeIMO.Drawing.HarfBuzz

`OfficeIMO.Drawing.HarfBuzz` is the optional full OpenType shaping adapter for
OfficeIMO renderers. It implements the shared
`IOfficeTextShapingProvider` contract with HarfBuzz GSUB/GPOS processing while
keeping `OfficeIMO.Drawing` and `OfficeIMO.Pdf` dependency-light.

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

The package uses HarfBuzzSharp `14.2.1.1` and matching Windows, Linux, macOS,
and WebAssembly native assets. HarfBuzzSharp is MIT licensed; applications
should include the upstream notices required by their own distribution policy.
