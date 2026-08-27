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
