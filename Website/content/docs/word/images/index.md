---
title: Images
description: Adding images to Word documents from files, streams, base64, and URLs with OfficeIMO.Word.
order: 13
---

# Images

OfficeIMO provides multiple ways to insert images into Word documents. Images can be added from local files, streams, base64-encoded strings, URLs, and embedded resources. You can control sizing, wrapping mode, cropping, rotation, and positioning.

## Adding an Image from a File

The simplest way to add an image is from a local file path:

```csharp
using OfficeIMO.Word;

using var document = WordDocument.Create("images.docx");

var paragraph = document.AddParagraph();
paragraph.AddImage("photo.jpg", width: 300, height: 200);

document.Save();
```

The `width` and `height` parameters are in pixels. If omitted, OfficeIMO uses the image's native dimensions.

## Adding an Image from a Stream

```csharp
using var imageStream = File.OpenRead("logo.png");

var paragraph = document.AddParagraph();
paragraph.AddImage(imageStream, "logo.png", width: 150, height: 50);
```

## Adding an Image from Base64

```csharp
string base64 = Convert.ToBase64String(File.ReadAllBytes("icon.png"));

var paragraph = document.AddParagraph();
paragraph.AddImageFromBase64(base64, "icon.png", width: 64, height: 64);
```

## Adding an Image from a URL

Download and embed an image directly from a URL:

```csharp
var image = document.AddImageFromUrl(
    "https://example.com/banner.jpg",
    width: 600,
    height: 200
);
```

## Adding an Image from an Embedded Resource

```csharp
var paragraph = document.AddParagraph();
paragraph.AddImageFromResource(
    typeof(MyClass).Assembly,
    "MyNamespace.Resources.logo.png",
    width: 200,
    height: 60
);
```

## Wrapping Modes

Control how text wraps around the image using the `WrapTextImage` enum:

```csharp
var paragraph = document.AddParagraph();
paragraph.AddImage("photo.jpg", 300, 200,
    wrapImageText: WrapTextImage.InLineWithText);
```

Available wrapping modes:

| Mode | Description |
|------|-------------|
| `WrapTextImage.InLineWithText` | Image sits inline with text (default) |
| `WrapTextImage.Square` | Text wraps around a square boundary |
| `WrapTextImage.Tight` | Text wraps tightly around the image shape |
| `WrapTextImage.Behind` | Image appears behind text |
| `WrapTextImage.InFront` | Image appears in front of text |
| `WrapTextImage.TopAndBottom` | Text appears above and below only |

## Image Positioning (Anchored Images)

When using a wrapping mode other than `InLineWithText`, the image is anchored and you can control its position:

```csharp
var paragraph = document.AddParagraph();
paragraph.AddImage("chart.png", 400, 300,
    wrapImageText: WrapTextImage.Square);

// Access the image through the paragraph
var image = document.Images.Last();
image.horizontalPosition.Offset = 914400;  // 1 inch from anchor in EMUs
```

## VML Images

For legacy compatibility, OfficeIMO supports VML (Vector Markup Language) images:

```csharp
// Add VML image to the document body
document.AddImageVml("watermark.png", width: 200, height: 100);

// Add VML image to a header
document.Sections[0].Header.Default.AddImageVml("header-logo.png", 150, 40);
```

## Image Properties

The `WordImage` class exposes properties for advanced image control:

```csharp
var image = document.Images[0];

// Cropping (values in percentage * 1000)
image.CropTop = 10000;      // crop 10% from top
image.CropBottom = 5000;    // crop 5% from bottom

// Effects
image.GrayScale = true;
image.LuminanceBrightness = 20000;    // brightness adjustment
image.LuminanceContrast = 10000;      // contrast adjustment

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
