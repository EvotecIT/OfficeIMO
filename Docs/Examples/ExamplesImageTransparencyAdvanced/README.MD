## Modifying transparency of an existing image

This advanced example loads a document, locates an image and updates its transparency.

```csharp
using (WordDocument document = WordDocument.Load(filePath)) {
    var image = document.Images[0];
    image.Transparency = 75; // make the image mostly transparent
    document.Save();
}
```
