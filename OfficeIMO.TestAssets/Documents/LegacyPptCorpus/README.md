# Legacy PowerPoint corpus

`BasicPowerPoint.ppt` was generated from the OfficeIMO PowerPoint basic example and converted to the
PowerPoint 97-2003 binary format with LibreOffice. It exercises positioned title, subtitle, and body
text shapes without embedding third-party content.

`PicturePowerPoint.ppt` was generated from a one-slide PPTX containing
`OfficeIMO.TestAssets/Images/EvotecLogo.png`, then converted to PowerPoint 97-2003 format with
LibreOffice. Its image is stored as a PNG BLIP in the compound file's `Pictures` delay stream. The
slide picture frame resolves that image through the document-level OfficeArt BStore and a one-based
`pib` property.

To regenerate a fixture from its PPTX source, use LibreOffice's headless converter:

```sh
soffice --headless --convert-to ppt --outdir <output-directory> <source>.pptx
```

Keep generated fixtures small and free of confidential or third-party content. Record the source
asset and the binary-format behavior covered by each new fixture here.
