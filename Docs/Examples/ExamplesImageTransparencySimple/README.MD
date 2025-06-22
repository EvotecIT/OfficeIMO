## Setting image transparency

This example demonstrates how to create a document with an image and apply a simple transparency value.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    var paragraph = document.AddParagraph();
    var image = paragraph.AddImage("image.png", 100, 100);
    image.Transparency = 30; // 30% transparent
    document.Save();
}
```
