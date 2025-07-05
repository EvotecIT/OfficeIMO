## Setting a background image

This example embeds an image behind the document content.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.Background.SetImage("image.png");
    document.Save();
}
```
