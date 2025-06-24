# Custom Paragraph Styles

OfficeIMO allows registering custom paragraph styles globally using `WordParagraphStyle.RegisterCustomStyle`.
Registered styles become available when setting the style identifier on a paragraph.

```csharp
var style = new Style { Type = StyleValues.Paragraph, StyleId = "MyStyle" };
WordParagraphStyle.RegisterCustomStyle("MyStyle", style);

using (WordDocument document = WordDocument.Create(filePath)) {
    document.AddParagraph("Hello world").SetStyleId("MyStyle");
    document.Save();
}
```
