---
title: Images
description: "Adding images to Word documents from files, streams, base64, and URLs with OfficeIMO.Word. Includes examples and API links."
order: 13
---

OfficeIMO provides multiple ways to insert images into Word documents. Images can be added from local files, streams, base64-encoded strings, URLs, and embedded resources. You can control sizing, wrapping mode, cropping, rotation, and positioning.

## Adding an Image from a File

The simplest way to add an image is from a local file path:

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using var document = WordDocument.Create("images.docx");

var paragraph = document.AddParagraph();
paragraph.AddImage("photo.jpg", width: OfficeLength.Pixels(300), height: OfficeLength.Pixels(200));

document.Save();
```

The `width` and `height` parameters take `OfficeLength`, so each value states its unit. If omitted, OfficeIMO uses the image's native dimensions.

## Adding an Image from a Stream

```csharp
using var imageStream = File.OpenRead("logo.png");

var paragraph = document.AddParagraph();
paragraph.AddImage(imageStream, "logo.png", width: OfficeLength.Pixels(150), height: OfficeLength.Pixels(50));
```

## Adding an Image from Base64

```csharp
string base64 = Convert.ToBase64String(File.ReadAllBytes("icon.png"));

var paragraph = document.AddParagraph();
using var decoded = new MemoryStream(Convert.FromBase64String(base64));
paragraph.AddImage(decoded, "icon.png", width: OfficeLength.Pixels(64), height: OfficeLength.Pixels(64));
```

## Adding an Image from a URL

Download and embed an image directly from a URL:

```csharp
var image = await document.AddImageAsync(
    new Uri("https://example.com/banner.jpg"),
    width: OfficeLength.Pixels(600),
    height: OfficeLength.Pixels(200)
);
```

## Adding an Image from an Embedded Resource

```csharp
var paragraph = document.AddParagraph();
using var resource = typeof(MyClass).Assembly.GetManifestResourceStream("MyNamespace.Resources.logo.png")
    ?? throw new InvalidOperationException("The embedded logo was not found.");
paragraph.AddImage(resource, "logo.png", OfficeLength.Pixels(200), OfficeLength.Pixels(60));
```

## Wrapping Modes

Control how text wraps around the image using the `WordImageTextWrapping` enum:

```csharp
var paragraph = document.AddParagraph();
paragraph.AddImage("photo.jpg", OfficeLength.Pixels(300), OfficeLength.Pixels(200),
    new WordImageOptions { Wrapping = WordImageTextWrapping.InLineWithText });
```

Available wrapping modes:

| Mode | Description |
|------|-------------|
| `WordImageTextWrapping.InLineWithText` | Image sits inline with text (default) |
| `WordImageTextWrapping.Square` | Text wraps around a square boundary |
| `WordImageTextWrapping.Tight` | Text wraps tightly around the image shape |
| `WordImageTextWrapping.BehindText` | Image appears behind text |
| `WordImageTextWrapping.InFrontOfText` | Image appears in front of text |
| `WordImageTextWrapping.TopAndBottom` | Text appears above and below only |

## Image Positioning (Anchored Images)

When using a wrapping mode other than `InLineWithText`, the image is anchored and you can control its position:

```csharp
var paragraph = document.AddParagraph();
paragraph.AddImage("chart.png", 400, 300,
    wrapImageText: WordImageTextWrapping.Square);

// Access the image through the paragraph
var image = document.Images.Last();
image.HorizontalPositionOffset = OfficeLength.Inches(1);
```

## VML Images

For legacy compatibility, OfficeIMO supports VML (Vector Markup Language) images:

```csharp
// Add VML image to the document body
document.AddImage("watermark.png", width: OfficeLength.Pixels(200), height: OfficeLength.Pixels(100),
    options: new WordImageOptions { Vml = true });

// Add VML image to a header
var (header, _) = document.AddHeadersAndFooters();
header.AddParagraph().AddImage("header-logo.png", OfficeLength.Points(150), OfficeLength.Points(40),
    options: new WordImageOptions { Vml = true });
```

## Image Properties

The `WordImage` class exposes properties for advanced image control:

```csharp
var image = document.Images[0];

// Select fractions of the source image while keeping the displayed frame.
image.CropTopRatio = 0.1;
image.CropBottomRatio = 0.05;

// Effects
image.GrayScale = true;
image.LuminanceBrightness = 20;       // increase brightness by 20%
image.LuminanceContrast = 10;         // increase contrast by 10%

// Lock aspect ratio
image.NoChangeAspect = true;

// Prevent operations
image.NoCrop = true;
image.NoResize = true;
image.NoRotation = true;
```

## Image Description (Alt Text)

Set alt text for accessibility:

```csharp
var paragraph = document.AddParagraph();
paragraph.AddImage("chart.png", 400, 300,
    description: "Q4 revenue chart showing 15% growth");
```

## Listing All Images

Access all images in the document:

```csharp
using var document = WordDocument.Load("report.docx");

foreach (var image in document.Images) {
    Console.WriteLine($"Image: {image.Width}x{image.Height}");
}
```
