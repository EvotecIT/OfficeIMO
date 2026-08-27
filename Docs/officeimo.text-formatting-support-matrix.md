# Text formatting support matrix

OfficeIMO keeps text formatting in the package that owns each document format and uses `OfficeIMO.Drawing` for reusable rendering behavior. This matrix distinguishes native document semantics from stored-text transformations and rendered approximations.

## Shared contracts

- `OfficeTextCaseTransformer` applies culture-aware `Uppercase`, `Lowercase`, `TitleCase`, `SentenceCase`, and `ToggleCase` transformations to stored text. Format-specific helpers preserve the surrounding run, paragraph, hyperlink, or shape formatting.
- `OfficeTextDecorationStyle` provides `Single`, `Double`, `Dotted`, `Dashed`, and `Wavy` underline or strikethrough patterns for Drawing, SVG, raster images, and PDF.
- `OfficeTextBaseline` provides normal, superscript, and subscript measurement and rendering. Script text uses a reduced effective size and participates in wrapping, fitting, backgrounds, rotation, and bounds calculations.
- Compatibility Boolean properties such as `Underline` and `Strikethrough` continue to map to a single solid line. The typed style properties are the source of truth when a caller selects another pattern.

## Native authoring and round-trip coverage

| Package or format | Family, size, color | Bold and italic | Underline | Strikethrough | Script | Casing |
| --- | --- | --- | --- | --- | --- | --- |
| Word (`.docx`) | Native | Native | 18 Word variants | Single and double | Subscript and superscript | Native all caps and small caps; all stored-text transforms |
| Excel (`.xlsx`) | Native cell and rich-run styles | Native | Single, double, single-accounting, and double-accounting | Native single | Subscript and superscript | Cell, range, and rich-run stored-text transforms |
| PowerPoint (`.pptx`) | Native run styles, including highlight | Native | DrawingML underline variants | Single and double | Percentage baseline, with subscript and superscript helpers | Native small caps and all caps; all stored-text transforms |
| RTF (`.rtf`) | Native | Native | RTF underline variants | Single and double | Subscript and superscript | Native caps and small caps; all stored-text transforms |
| OpenDocument text and presentation (`.odt`, `.odp`) | Native | Native | Solid, dotted, dash, long-dash, dot-dash, dot-dot-dash, and wave; single or double | Same line styles and counts | Subscript and superscript | Native uppercase, lowercase, and capitalize; all stored-text transforms |
| OpenDocument spreadsheet (`.ods`) | Native cell styles | Native | Solid, dotted, dash, long-dash, dot-dash, dot-dot-dash, and wave; single or double | Same line styles and counts | Subscript and superscript | Native uppercase, lowercase, capitalize, and small caps |
| PDF | Authored into page content | Native standard or registered embedded faces | Single, double, dotted, dashed, and wavy | Single, double, dotted, dashed, and wavy | Subscript and superscript | Stored-text transforms |
| OneNote (`.one`) | Native | Native | Native single | Native single | Native subscript and superscript | Stored-text transforms |
| Visio (`.vsdx` family) | Native whole-text-block Character cells | Native | Native single and double | Native single and double | Native subscript and superscript | Native all caps, initial caps, and small caps; all stored-text transforms |

The Word, RTF, PowerPoint, OpenDocument, OneNote, and Visio rows describe typed read/write behavior, not only export appearance. Excel exposes the same native font model through rich runs and the common cell, range, and sheet helpers.

## Rendering coverage

| Surface | Behavior |
| --- | --- |
| Shared Drawing rich text | Family, size, color, background, bold, italic, all shared decoration patterns, and script-aware measurement/layout |
| SVG | Emits font, fill, weight, style, decoration, decoration style, reduced script size, and baseline shift. When underline and strike use different patterns on the same SVG text element, SVG applies one decoration style to both; the underline pattern wins. |
| Raster (`PNG`, `JPEG`, `TIFF`, `WebP`) | Draws shared decoration patterns and script baselines through the common raster canvas, including rotation and mirroring. These are the dependency-free export formats exposed by Word, Excel, PowerPoint, and Visio. |
| PDF | Draws shared decoration patterns as PDF graphics operators and includes double/wavy extents in page-text bounds |
| OneNote image/PDF export | Projects native OneNote run styles into shared Drawing, including native strike and script flags |
| Visio SVG/image export | Projects native Visio underline, strike, script, case, and small-cap semantics. Raster small caps use an uppercase approximation because the dependency-free raster font path has no small-cap glyph substitution. |

## Syntax-oriented formats

HTML keeps CSS text semantics in `OfficeIMO.Html`, including font properties, colors, text decoration, vertical alignment, and text transformation. Markdown, AsciiDoc, and LaTeX retain the emphasis, strike, subscript, superscript, or inline styling that their supported syntax profiles can represent; they do not invent font-family or point-size syntax where the target language has no portable equivalent. Conversion diagnostics remain the authority when a source style cannot be represented by the target syntax.

The Word↔ODT, PowerPoint↔ODP, and Excel↔ODS adapters preserve the destination format's native underline counts or patterns, strike counts, and script placement where an exact representation exists. Excel accounting underlines, ODF decoration patterns unsupported by Excel, and ODF small caps in Excel are reported as approximations or unsupported semantics instead of being silently presented as exact round trips.

## Conversion graph

Text formatting is verified at the reusable owner and at multi-hop boundaries. The contract depends on the destination:

- Editable Office targets retain native family, size, color, bold, italic, underline, strike, and script semantics where that target has an equivalent. Converter metadata retains richer source variants, such as Word wavy-double underline, Excel accounting underline, and PowerPoint capitalization or baseline percentage, across OfficeIMO-generated HTML round trips.
- Semantic HTML retains CSS typography and native converter metadata. Word, Excel, PowerPoint, and OneNote HTML adapters restore their representable native run styles. Managed HTML rendering applies the same style model to PDF, SVG, and raster output.
- Markdown, AsciiDoc, and LaTeX preserve only formatting represented by their supported syntax profiles. Font family, point size, arbitrary color, and underline variants are intentionally diagnosed rather than encoded as nonportable syntax.
- PDF output preserves the rendered appearance and text-run styling supported by the managed PDF engine. Importing PDF back into Word, Excel, PowerPoint, HTML, RTF, or OpenDocument is reconstruction from PDF logical or positioned content; it cannot recover source-only font semantics that were not encoded in the PDF.
- Image output is deliberately flattened. PNG, JPEG, TIFF, and WebP preserve pixels; SVG preserves the emitted vector text and decoration attributes. Image conversions do not claim editable font semantics.

| Conversion family | Text-formatting contract |
| --- | --- |
| Word ↔ HTML | Native family, size, color, bold, italic, highlight, underline variants, single/double strike, scripts, and native caps metadata where representable |
| Excel ↔ HTML | Cell and rich-run family, size, color, bold, italic, accounting/single/double underline metadata, strike, and scripts |
| PowerPoint ↔ HTML | Per-run family, size, color, bold, italic, DrawingML underline/strike variants, capitalization, and baseline percentage metadata |
| OneNote ↔ HTML | Native family, size, foreground/highlight colors, bold, italic, underline, strike, and scripts; OneNote has Boolean rather than patterned decorations |
| Word ↔ RTF and RTF ↔ HTML | Native RTF character formatting and supported decoration variants, scripts, caps, family, size, and color |
| Word ↔ ODT, Excel ↔ ODS, PowerPoint ↔ ODP | Exact native mappings where available, with explicit approximation or omission diagnostics for target gaps |
| Word/Excel/PowerPoint/HTML/RTF/Markdown/AsciiDoc/LaTeX/OneNote/ODT/ODS/ODP/MHTML/Visio → PDF | Fixed-layout style projection through the owning native or semantic adapter and shared PDF text model |
| PDF → Word/Excel/PowerPoint/HTML/RTF/ODT/ODS/ODP | Bounded editable or semantic reconstruction; visual fidelity and recoverable text properties depend on PDF content |
| Word/Excel/PowerPoint/HTML/OneNote/Visio/email/EPUB/ODT/ODS/ODP/PDF → image | All five shared formats: PNG, SVG, JPEG, TIFF, and WebP; visual style retention, not editable typography |
| Word ↔ Google Docs | Native family, size, color, bold, italic, single underline/strike, highlight, small caps, superscript, and subscript; all-caps is materialized and richer Word variants are diagnosed or handled by Drive fallback |
| Excel ↔ Google Sheets | Native cell and rich-run family, size, color, bold, italic, single underline, and strike; the Sheets API has no script, casing, or underline-variant fields |
| PowerPoint ↔ Google Slides | Native per-run family, size, color, bold, italic, single underline/strike, small caps, superscript, and subscript; all-caps is materialized and richer DrawingML variants use their closest supported appearance |
| ADF ↔ Markdown/HTML and Confluence bodies | Strong, emphasis, strike, underline, subscript, and superscript survive the supported syntax pipeline; ADF text/background colors and richer marks without a portable Markdown equivalent remain diagnosed limitations |
| OfficeIMO Markup → Word/Excel/PowerPoint | Block- or target-range family, size, color/highlight, bold, italic, underline variants, strike, scripts, case transforms, and small caps where the destination has a native representation; Excel additionally accepts accounting underline variants |
| CSV ↔ Excel | Values and records only. CSV/TSV has no font, decoration, script, case-metadata, formula, drawing, layout, or multi-sheet model, so typography is intentionally not a portable contract |

`OfficeConversionCapabilityCatalog` exposes every image source/format pair; authenticated Google Docs, Sheets, and Slides import/export pairs; ADF and Confluence content projections; CSV/Excel interchange; and OfficeIMO Markup authoring routes, in addition to the local document, semantic, and PDF routes. Remote imports are marked `RemoteResource` because they accept a service document identifier, while materialized Confluence pages are marked `ObjectModel`. Each route also declares a machine-readable typography contract: editable equivalent, semantic equivalent, syntax subset, fixed-layout appearance, PDF reconstruction, vector appearance, flattened raster, or data-only. This makes PDF→SVG/JPEG/TIFF/WebP and the Word, Excel, PowerPoint, OneNote, Visio, HTML, email, EPUB, OpenDocument, Google Workspace, ADF/Confluence, CSV, and OfficeIMO Markup paths discoverable without overstating editable fidelity.

The chain-level regression suite includes HTML→PDF→all image formats, Word→PDF→all image formats, and MHTML→PDF→all image formats in addition to direct source exports. Runtime tests exercise all five image formats for Word, Excel, PowerPoint, HTML, OneNote, Visio, email, EPUB, ODT, ODS, ODP, and PDF. Source-specific tests verify that native and semantic styles reach the shared Drawing/PDF owners before those common encoders run.

## Usage

Shared rendering and PDF use the same decoration and baseline enums:

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

PdfDocument document = PdfDocument.Create(pdf => pdf.Content(content => content
    .Paragraph(paragraph => paragraph
        .Underlined("double", OfficeTextDecorationStyle.Double)
        .Text(" H")
        .Subscript("2")
        .Text("O ")
        .Strikethrough("retired", OfficeTextDecorationStyle.Wavy))));
```

Visio maps the supported subset to native ShapeSheet Character cells:

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Visio;

VisioShape shape = page.AddRectangle(2, 4, 2.5, 1, "Release Status");
shape.TextStyle = new VisioTextStyle {
    FontFamily = "Aptos",
    Size = 14,
    Color = OfficeColor.DarkBlue,
    Bold = true,
    UnderlineStyle = OfficeTextDecorationStyle.Double,
    StrikethroughStyle = OfficeTextDecorationStyle.Single,
    Baseline = OfficeTextBaseline.Superscript,
    Capitalization = VisioTextCapitalization.AllCaps
};
shape.TransformTextCase(OfficeTextCase.TitleCase);
```

Format-native display casing and stored-text casing are intentionally separate. Display casing leaves the underlying text unchanged when the file format supports it; `TransformTextCase` changes the stored text and preserves the existing formatting object.

## Compatibility boundaries

- Existing Boolean underline and strike APIs remain source compatible and mean one solid line.
- New optional constructor parameters were appended to shared rich-text and PDF run constructors so existing positional calls keep their meaning.
- Visio supports only none, single, and double decoration lines natively. Assigning dotted, dashed, or wavy patterns to a `VisioTextStyle` is rejected instead of silently flattening the request.
- OneNote stores Boolean underline and strike properties; decoration variants are not claimed as native OneNote features.
- Conversion packages may approximate or report unsupported source formatting. They should not silently advertise the target artifact as native-equivalent when its format cannot carry the source style.
