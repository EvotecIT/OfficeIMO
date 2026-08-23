# OfficeIMO.Drawing.SixLabors

`OfficeIMO.Drawing.SixLabors` is the optional managed font engine for OfficeIMO
renderers on .NET 8 and newer. It adds WOFF2, CFF/CFF2, variable-font instances,
OpenType shaping, bidirectional layout, and outline rendering without adding
those dependencies to `OfficeIMO.Core`.

```csharp
using OfficeIMO.Drawing.SixLabors;
using OfficeIMO.Html;

var options = new HtmlRenderOptions()
    .UseSixLaborsFontPrograms();
```

Fallback packs snapshot their faces, provider-selected variable instances,
Unicode ranges, and fallback order under a deterministic SHA-256 fingerprint:

```csharp
var fonts = new OfficeFontFaceCollection()
    .UseSixLaborsFontPrograms()
    .Add("Latin", latinWoff2)
    .Add("Arabic", arabicTtf)
    .AddFallbackFamily("Arabic");

var pack = new OfficeFontFallbackPack("report-fonts-v1", "Latin, Arabic", fonts);
var options = new HtmlRenderOptions().UseFontFallbackPack(pack);
```

The provider is opt-in. Built-in TrueType and WOFF 1 behavior remains available
without this package. Static TrueType and CFF1 sfnt programs can be embedded as
PDF fonts. WOFF2, CFF2, and active variable-font instances are exposed as
outline programs; PDF conversion can preserve their appearance using accessible
vector text rather than claiming that the original container is a static
embeddable font.

The provider implements `IOfficeBoundedFontProgram`. Outline expansion checks
cancellation while glyph contours are produced and stops before exceeding the
renderer-supplied point budget. Coverage planning treats variation selectors,
join controls, and bidi controls as shaping inputs rather than requiring invalid
standalone glyphs, so emoji variation and ZWJ sequences stay with the selected
emoji face.

This package depends on SixLabors.Fonts 3.1.0, which validates a Six Labors
license at compile time. Supply `SixLaborsLicenseKey`, `SixLaborsLicenseFile`, or
a discovered `sixlabors.lic` file as described by Six Labors. This source tree
also maps the `SIXLABORS_LICENSE_KEY` environment variable to the MSBuild
property for automated builds. Never commit the supplied license file or value.
Open-source maintainers can apply through the
[Six Labors licensing portal](https://licensing.sixlabors.com/); confirm the
applicable terms for your own build and distribution scenario. See
`THIRD-PARTY-NOTICES.md` in the package.
