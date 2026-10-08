# ICC image fixtures

`generate.py` encodes constant-color PNG, JPEG and classic TIFF images with Pillow
and records media-relative LittleCMS reference channels in `expected.csv`.
Fixtures cover associated and embedded profiles, grayscale, RGB, CMYK and alpha.
Nonuniform baseline/progressive JPEGs also carry EXIF orientation 6.
JPEG expectations use the decoded JPEG samples, accounting for lossy encoding.

The RGB and CMYK profiles and their licenses/provenance are maintained in
`OfficeIMO.Drawing.Tests/TestAssets/IccColorCorpus`. The gray gamma-1.8 profile is
generated through LittleCMS; it is not copied from an external profile library.
`oversized-profile.png` contains a compressed profile exceeding the 4 MiB limit.

Generation environment: Pillow 11.3.0, LittleCMS 2.17.
These tools are used only to regenerate test fixtures. They are not OfficeIMO
runtime or normal build requirements.
