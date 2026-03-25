---
title: Slides
description: Creating PowerPoint slides with text boxes, shapes, images, bullets, and charts.
order: 31
---

# Slides

This guide covers creating and populating PowerPoint slides using the Open XML SDK through OfficeIMO. Slides can contain text boxes, bulleted lists, shapes, images, and embedded charts.

## Creating a Slide

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

using var pptx = PresentationDocument.Create("presentation.pptx",
    PresentationDocumentType.Presentation, true);

var presentationPart = pptx.AddPresentationPart();
presentationPart.Presentation = new Presentation(
    new SlideIdList(),
    new SlideSize { Cx = 9144000, Cy = 6858000 },  // Standard 10x7.5 inches
    new NotesSize { Cx = 6858000, Cy = 9144000 }
);

// Add a slide
var slidePart = presentationPart.AddNewPart<SlidePart>();
slidePart.Slide = new Slide(new CommonSlideData(new ShapeTree()));

// Register the slide in the presentation
var slideIdList = presentationPart.Presentation.SlideIdList!;
slideIdList.Append(new SlideId {
    Id = 256,
    RelationshipId = presentationPart.GetIdOfPart(slidePart)
});

pptx.Save();
```

## Adding a Text Box

```csharp
var shapeTree = slidePart.Slide.CommonSlideData!.ShapeTree!;

var textBox = new Shape(
    new NonVisualShapeProperties(
        new NonVisualDrawingProperties { Id = 2, Name = "TextBox 1" },
        new NonVisualShapeDrawingProperties(),
        new ApplicationNonVisualDrawingProperties()
    ),
    new ShapeProperties(
        new D.Transform2D(
            new D.Offset { X = 1000000, Y = 1000000 },     // Position in EMUs
            new D.Extents { Cx = 7000000, Cy = 1500000 }   // Size in EMUs
        ),
        new D.PresetGeometry(new D.AdjustValueList()) { Preset = D.ShapeTypeValues.Rectangle }
    ),
    new TextBody(
        new D.BodyProperties(),
        new D.Paragraph(
            new D.Run(
                new D.RunProperties { Language = "en-US", FontSize = 2400, Bold = true },
                new D.Text("Quarterly Results")
            )
        )
    )
);

shapeTree.Append(textBox);
```

## Bulleted Lists

```csharp
var bulletBody = new TextBody(
    new D.BodyProperties(),
    new D.ListStyle(),
    // Bullet 1
    new D.Paragraph(
        new D.ParagraphProperties(
            new D.BulletFont { Typeface = "Arial" },
            new D.CharacterBullet { Char = "\u2022" }  // bullet character
        ),
        new D.Run(
            new D.RunProperties { Language = "en-US", FontSize = 1800 },
            new D.Text("Revenue grew 15% year-over-year")
        )
    ),
    // Bullet 2
    new D.Paragraph(
        new D.ParagraphProperties(
            new D.BulletFont { Typeface = "Arial" },
            new D.CharacterBullet { Char = "\u2022" }
        ),
        new D.Run(
            new D.RunProperties { Language = "en-US", FontSize = 1800 },
            new D.Text("Customer base expanded to 10,000+")
        )
    ),
    // Bullet 3
    new D.Paragraph(
        new D.ParagraphProperties(
            new D.BulletFont { Typeface = "Arial" },
            new D.CharacterBullet { Char = "\u2022" }
        ),
        new D.Run(
            new D.RunProperties { Language = "en-US", FontSize = 1800 },
            new D.Text("New product line launching Q2")
        )
    )
);
```

## Adding Images

Embed an image into a slide:

```csharp
// Add the image part
var imagePart = slidePart.AddImagePart(ImagePartType.Png);
using (var imageStream = File.OpenRead("chart.png")) {
    imagePart.FeedData(imageStream);
}

var imageRelId = slidePart.GetIdOfPart(imagePart);

// Create the picture shape
var picture = new Picture(
    new NonVisualPictureProperties(
        new NonVisualDrawingProperties { Id = 3, Name = "Picture 1" },
        new NonVisualPictureDrawingProperties(
            new D.PictureLocks { NoChangeAspect = true }
        ),
        new ApplicationNonVisualDrawingProperties()
    ),
    new BlipFill(
        new D.Blip { Embed = imageRelId },
        new D.Stretch(new D.FillRectangle())
    ),
    new ShapeProperties(
        new D.Transform2D(
            new D.Offset { X = 1500000, Y = 2500000 },
            new D.Extents { Cx = 6000000, Cy = 4000000 }
        ),
        new D.PresetGeometry(new D.AdjustValueList()) { Preset = D.ShapeTypeValues.Rectangle }
    )
);

shapeTree.Append(picture);
```

## Shapes

Add geometric shapes (rectangles, ellipses, arrows, etc.):

```csharp
var rectangle = new Shape(
    new NonVisualShapeProperties(
        new NonVisualDrawingProperties { Id = 4, Name = "Rectangle 1" },
        new NonVisualShapeDrawingProperties(),
        new ApplicationNonVisualDrawingProperties()
    ),
    new ShapeProperties(
        new D.Transform2D(
            new D.Offset { X = 500000, Y = 5000000 },
            new D.Extents { Cx = 2000000, Cy = 800000 }
        ),
        new D.PresetGeometry(new D.AdjustValueList()) { Preset = D.ShapeTypeValues.RoundRectangle },
        new D.SolidFill(new D.RgbColorModelHex { Val = "4472C4" })
    ),
    new TextBody(
        new D.BodyProperties { Anchor = D.TextAnchoringTypeValues.Center },
        new D.Paragraph(
            new D.ParagraphProperties { Alignment = D.TextAlignmentTypeValues.Center },
            new D.Run(
                new D.RunProperties { Language = "en-US", FontSize = 1400 },
                new D.Text("Click Here")
            )
        )
    )
);

shapeTree.Append(rectangle);
```

## Charts

For embedded charts, the recommended approach is to create the chart using OfficeIMO.Excel, export it as an image, and embed it into the slide as shown in the "Adding Images" section above. This avoids the complexity of the Open XML chart markup while producing visually identical results.

## Tips

- All positions and sizes in Open XML are specified in **EMUs** (English Metric Units). 1 inch = 914400 EMUs.
- Standard slide dimensions: 9144000 x 6858000 EMUs (10 x 7.5 inches).
- Use the `NonVisualDrawingProperties.Id` to assign unique IDs to each shape on a slide.
- Shape types are defined in `D.ShapeTypeValues` and include `Rectangle`, `Ellipse`, `RightArrow`, `Star5`, and many more.
