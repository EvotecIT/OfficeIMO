## Drawing VML Shapes

`WordShape` exposes factory methods for creating simple VML shapes directly on a paragraph.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    var paragraph = document.AddParagraph("Shapes demo");
    // rectangle
    paragraph.AddShape(100, 50, SixLabors.ImageSharp.Color.Red);
    // ellipse
    WordShape.AddEllipse(paragraph, 80, 40, SixLabors.ImageSharp.Color.Green);
    // line
    WordShape.AddLine(paragraph, 0, 60, 100, 60, SixLabors.ImageSharp.Color.Blue, 2);
    // polygon
    WordShape.AddPolygon(paragraph, "0,0 50,0 50,50 0,50", SixLabors.ImageSharp.Color.Yellow, SixLabors.ImageSharp.Color.Black);
    document.Save();
}
```
