Consolidated HTML conversion examples

0. Html00_AllInOne – Use comprehensive input (all.html) → DOCX → round-trip HTML.
1. Html01_LoadAndRoundTripBasics – Load basic HTML → DOCX, round-trip to HTML.
2. Html02_SaveAsHtmlFromWord – Build a Word document programmatically and export to HTML with options.
3. Html03_TextFormatting – Inline styles and tags: underline/strike, sup/sub, transforms, spacing.
4. Html04_ListsAndNumbering – Nested lists, custom starts, roman/alpha types.
5. Html05_TablesComplex – thead/tbody/tfoot, caption, rowspan/colspan.
6. Html06_ImagesAllModes – Data URI, relative, absolute image references.
7. Html07_LinksAndAnchors – External links and internal anchors/ids.
8. Html08_SemanticsAndCitations – blockquote/cite, abbr/dfn, figure/figcaption.
9. Html09_CodePreWhitespace – code vs pre, whitespace handling.
10. Html10_OptionsAndAsync – Additional head tags, style options, async round-trip.

All examples share the signature: Example(string folderPath, bool openWord)
and write outputs to the runner’s Documents folder.
