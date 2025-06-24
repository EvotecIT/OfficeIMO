## Custom Paragraph Styles

Custom paragraph styles can be registered globally before creating a document.

```csharp
var custom = new Style { Type = StyleValues.Paragraph, StyleId = "MyStyle" };
WordParagraphStyle.RegisterCustomStyle("MyStyle", custom);

using (WordDocument document = WordDocument.Create(filePath)) {
    document.AddParagraph("Hello").SetStyleId("MyStyle");
    document.Save();
}
```

Built-in styles may also be replaced:

```csharp
var normal = new Style { Type = StyleValues.Paragraph, StyleId = "Normal" };
WordParagraphStyle.OverrideBuiltInStyle(WordParagraphStyles.Normal, normal);
```
