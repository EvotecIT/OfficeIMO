# Drawing VML Shapes

`WordShape` exposes factory methods for creating simple VML shapes directly on a paragraph.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    var paragraph = document.AddParagraph("Shapes demo");
    // rectangle
    paragraph.AddShape(OfficeIMO.OfficeLength.Points(100), OfficeIMO.OfficeLength.Points(50), OfficeIMO.OfficeColor.Red);
    // ellipse
    WordShape.AddEllipse(paragraph, OfficeIMO.OfficeLength.Points(80), OfficeIMO.OfficeLength.Points(40), OfficeIMO.OfficeColor.Green);
    // line
    WordShape.AddLine(paragraph, OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(60), OfficeIMO.OfficeLength.Points(100), OfficeIMO.OfficeLength.Points(60), OfficeIMO.OfficeColor.Blue, OfficeIMO.OfficeLength.Points(2));
    // polygon
    WordShape.AddPolygon(paragraph, "0,0 50,0 50,50 0,50", OfficeIMO.OfficeColor.Yellow, OfficeIMO.OfficeColor.Black);
    document.Save();
}
```
