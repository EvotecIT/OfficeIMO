# HTML conversion examples

This folder contains runnable HTML-to-Office and managed HTML-to-PDF examples.

## Run

```powershell
dotnet run --project ..\..\OfficeIMO.Examples.csproj -f net10.0
```

## Examples

- `Html00_AllInOne`: comprehensive input (`all.html`) to DOCX and round-trip HTML.
- `Html01_LoadAndRoundTripBasics`: basic HTML to DOCX and back to HTML.
- `Html02_SaveAsHtmlFromWord`: build a Word document and export to HTML with options.
- `Html03_TextFormatting`: inline styles and tags.
- `Html04_ListsAndNumbering`: nested lists, custom starts, roman and alpha list types.
- `Html05_TablesComplex`: table sections, captions, rowspan, and colspan.
- `Html06_ImagesAllModes`: data URI, relative, and absolute images.
- `Html07_LinksAndAnchors`: external links and internal anchors.
- `Html08_SemanticsAndCitations`: semantic HTML and citations.
- `Html09_CodePreWhitespace`: code, preformatted text, and whitespace handling.
- `Html10_OptionsAndAsync`: additional head tags, style options, and async round-trip.
- `HtmlAccessibleReport`: a two-page tagged PDF report with embedded fonts, source-aware validation, table headers, lists, a described figure, decorative SVG, internal navigation and a labelled form control. Run `dotnet run --project OfficeIMO.Examples/OfficeIMO.Examples.csproj -f net10.0 -- --html-accessible` from the repository root. The [HTML/PDF guide](../../../OfficeIMO.Html/PDF.md#accessibility-validation) describes independent validation and reading review.

The examples write outputs to the runner's `Documents` folder. The accessible report opens generated PDFs only when `--open-pdf` is supplied.
