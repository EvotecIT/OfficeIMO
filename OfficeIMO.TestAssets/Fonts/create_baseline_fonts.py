"""Regenerate the small, renamed OFL font fixtures used by visual tests."""
from pathlib import Path
from fontTools import subset
from fontTools.ttLib import TTFont

output = Path(__file__).resolve().parent
source = output.parents[1] / "Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts"
characters = set(range(32, 384)) | {0x2022, 0x2026, 0x20AC, 0x2713, 0x2714, 0x2717}
for style in ("Regular", "Bold", "Italic", "BoldItalic"):
    font = TTFont(source / f"Carlito-{style}.ttf", recalcTimestamp=False)
    options = subset.Options()
    options.name_IDs = ["*"]
    options.name_legacy = True
    options.name_languages = ["*"]
    subsetter = subset.Subsetter(options=options)
    subsetter.populate(unicodes=characters)
    subsetter.subset(font)
    # The original font's reserved name must not identify modified font files.
    for record in font["name"].names:
        if record.nameID in (1, 3, 4, 6, 16, 17):
            value = record.toUnicode().replace("Carlito", "OfficeIMO Baseline Sans")
            if record.nameID == 6:
                value = "OfficeIMOBaselineSans-" + style
            record.string = value.encode(record.getEncoding())
    font.save(output / f"OfficeIMOBaselineSans-{style}.ttf")
