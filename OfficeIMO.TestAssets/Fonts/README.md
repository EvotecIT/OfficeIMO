# Visual baseline fonts

OfficeIMO Baseline Sans is a small Latin subset of the repository's Carlito fonts,
renamed to respect the reserved font name. The original copyright and SIL Open
Font License apply; see `OFL-Carlito.txt`.

These fixtures keep spreadsheet typography independent of installed system fonts.
They include regular, bold, italic, and bold italic faces. They are test assets,
not a replacement font for document exports.

Regenerate with Python and fontTools 4.46.0:

```sh
python OfficeIMO.TestAssets/Fonts/create_baseline_fonts.py
```
