# Legacy PowerPoint corpus

`BasicPowerPoint.ppt` was generated from the OfficeIMO PowerPoint basic example and converted to the
PowerPoint 97-2003 binary format with LibreOffice. It exercises positioned title, subtitle, and body
text shapes without embedding third-party content.

`PicturePowerPoint.ppt` was generated from a one-slide PPTX containing
`OfficeIMO.TestAssets/Images/EvotecLogo.png`, then converted to PowerPoint 97-2003 format with
LibreOffice. Its image is stored as a PNG BLIP in the compound file's `Pictures` delay stream. The
slide picture frame resolves that image through the document-level OfficeArt BStore and a one-based
`pib` property.

`ShapePowerPoint.ppt` was generated from a one-slide PPTX built with `python-pptx`, then converted
with LibreOffice. It contains text-bearing and empty preset shapes from several geometry families,
straight, bent, and curved connectors, plus a group containing a rounded rectangle and an ellipse.
The fixture exercises OfficeArt preset mapping, native connector projection, nested group parsing,
child coordinate systems, exact no-op saves, and incremental connector and outer-group geometry
edits. Its SHA-256 digest is
`50f5094a5004cba187defe96d9d15f7e7302fcd8c1176711a1538eeee05eae20`.

`TransformPowerPoint.ppt` was generated from a one-slide PPTX built with `python-pptx`, then
converted with LibreOffice. It contains clockwise and counterclockwise shape rotations, horizontal
and vertical mirroring, plus a nested group. The fixture exercises OfficeArt fixed-point rotation,
FSP flip flags, exact no-op saves, and loss blocking for unsupported transform edits. Its SHA-256 digest is
`eef8191b166d1a2eb4714bddac250390075eb16122c7ecbce6dd285ef707d2e2`.

To regenerate a fixture from its PPTX source, use LibreOffice's headless converter:

```sh
soffice --headless --convert-to ppt --outdir <output-directory> <source>.pptx
```

Keep generated fixtures small and free of confidential or third-party content. Record the source
asset and the binary-format behavior covered by each new fixture here.
