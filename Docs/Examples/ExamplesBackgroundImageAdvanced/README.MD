## Background image from stream with custom size

This advanced example loads the image from a stream and specifies the width and height.

```csharp
using var stream = File.OpenRead("image.png");
using (WordDocument document = WordDocument.Create(filePath)) {
    document.Background.SetImage(stream, "image.png", 600, 800);
    document.Save();
}
```
