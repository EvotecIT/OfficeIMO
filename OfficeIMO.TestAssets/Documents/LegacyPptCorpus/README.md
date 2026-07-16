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

`ConnectedPowerPoint.ppt` was generated from a one-slide PPTX built with `python-pptx`, with an
Open XML connector explicitly attached between a rounded rectangle and a diamond, then converted
with LibreOffice. It exercises OfficeArt connector solver rules, start/end shape identifiers,
connection-site indexes, native DrawingML attachment projection, and exact no-op saves. Its
SHA-256 digest is
`9ec3ea4c6f7a0d6fca06a00a269cbe15b66ded0f4a124e9e5f029c75e3de3ef8`.

`AdjustedShapesPowerPoint.ppt` was generated from a one-slide PPTX built with `python-pptx`, with
explicit adjustments on a rounded rectangle, chevron, right arrow, donut, trapezoid, and arc, then
converted with LibreOffice. It exercises all signed OfficeArt adjustment slots, conservative
shape-family mapping, exact round-rectangle and donut guide projection, deliberate retention of
unmapped shape-specific values, and exact no-op saves. Its SHA-256 digest is
`cfbcdc8249b0cd886a6abcf6c00fe00eb85753eaa693dd933299d47b68a42599`.

`CroppedPicturePowerPoint.ppt` was generated from a one-slide PPTX containing two copies of a
four-quadrant PNG, then converted with LibreOffice. One picture uses four positive crop edges and
the other uses negative top and left crop-out values. It exercises signed 16.16 OfficeArt crop
decoding, native DrawingML source-rectangle projection, and exact no-op saves. Its SHA-256 digest is
`42b9007c1d995ecd0471bdc195a9b2a72acdc89dfbd4cee73f47000099068dc9`.

`PictureEffectsPowerPoint.ppt` was generated from a one-slide PPTX containing brightness,
negative and positive contrast, grayscale, bi-level, and recolor examples, then converted with
LibreOffice. It exercises signed effect values, the OfficeArt Boolean use/value masks, native
DrawingML luminance and monochrome projection, baked-image fallback, and exact no-op saves. Its
SHA-256 digest is
`db2eadf76110641fe46230949f04ac4c54cb158678b7cd301ceea193161e27d2`.

`ShadowPowerPoint.ppt` was generated from a one-slide PPTX containing two rounded rectangles with
45-degree and 135-degree outer shadows, then converted with LibreOffice. It exercises primary RGB
color, opacity, positive and negative signed EMU offsets, polar DrawingML projection, explicit
shadow visibility, and exact no-op saves. Its SHA-256 digest is
`33eff835786f06fd721f9f3b2b15300ea2d8aa96ec1dd79160c6e0811d207897`.

`AccessibilityPowerPoint.ppt` was authored as a one-slide PPTX and saved to PowerPoint 97-2003 by
Microsoft PowerPoint for Mac. Its rounded rectangle, picture, and connector have distinct object
names and descriptions. It exercises OfficeArt `wzName` and `wzDescription` decoding, native
non-visual metadata projection, Microsoft's four-byte CurrentUserAtom length overstatement, exact
no-op saves, and loss blocking for unsupported metadata edits. Its SHA-256 digest is
`ce8b78712e423238d9082503842152c07ebfc6770ebfc3654c0865b725f4e175`.

`RichTextPowerPoint.ppt` was authored as a one-slide PPTX and saved to PowerPoint 97-2003 by
Microsoft PowerPoint for Mac. Its text box contains separate bold red, italic green, underlined
blue, separator, and second-paragraph runs with sizes from 20 to 32 points. It exercises bounded
`StyleTextPropAtom` paragraph and character arrays, native emphasis, size, and direct RGB color
projection, document font-collection resolution, exact no-op saves, formatting-preserving geometry edits, and loss blocking for
unsupported formatting edits. Its SHA-256 digest is
`91d0c2a9ee4380357befeee9803d45b4bcdfb2774d84ea5360e3f8b3964dab88`.

`apache-poi-testdata/bug61881.ppt` is a public Apache POI slideshow fixture
that contains a PPT9 `BlipCollection9Container`, a PNG picture-bullet entity,
and a paragraph that references the entity by its zero-based picture-bullet
index. It exercises independent binary decoding and native DrawingML
`a:buBlip` projection. Its provenance and license are recorded alongside the
fixture.

To regenerate a fixture from its PPTX source, use LibreOffice's headless converter:

```sh
soffice --headless --convert-to ppt --outdir <output-directory> <source>.pptx
```

Keep generated fixtures small and free of confidential or third-party content. Record the source
asset and the binary-format behavior covered by each new fixture here.
