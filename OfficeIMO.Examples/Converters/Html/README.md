# HTML conversion examples

This folder contains runnable HTML conversion examples for `OfficeIMO.Word.Html`.

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

All examples share the `Example(string folderPath, bool openWord)` signature and write outputs to the runner's Documents folder.
