# Native integer image defaults

`generate.py` encodes eight synthetic 16-by-8 images with Pillow 12.3.0. It is an
optional fixture regeneration tool, not a runtime or ordinary test dependency.
Run it with Python and Pillow from this directory. `expected.csv` contains the
source sample values, JPEG decoded sample values, and native alpha policy;
`manifest.json` records SHA-256 hashes.

TIFF fixtures exercise RGB, gray, alpha, calibration tags, display orientation,
an ignored unspecified extra sample, and an embedded profile. JPEG fixtures
exercise gray/RGB calibration metadata. The embedded profile is the existing
independently generated LittleCMS RGB matrix fixture. Source image pixels and
the generator are repository-authored test data.

ECMA-388 15.3.7/M8.30 defines integer sRGB/gray defaults without a usable ICC
profile. Section 9.1.5.3/M2.26 excludes TIFF orientation and M2.83 ignores an
extra sample declared as unspecified. GhostXPS agrees with the raw TIFF viewbox
but changes sample order for unspecified extra channels in these fixtures; that
disagreement is retained separately from managed and
independent SVG/PDF color proof.
