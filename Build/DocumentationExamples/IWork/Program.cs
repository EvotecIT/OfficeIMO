#region QuickStart
using OfficeIMO.IWork;

var presentation = IWorkKeynoteDocument.Create(); // 960 x 540 points
var slide = presentation.AddSlide("E6F2FF");
slide.AddText("Quarterly review", 60, 100, 840, 120,
    fontName: "Arial", fontSizePoints: 40, color: "003366");
slide.AddRichText(new[] {
    new IWorkTextRunData("Revenue ", bold: true),
    new IWorkTextRunData("up 12%"),
    new IWorkTextRunData(", two open risks", underline: true)
}, 60, 300, 840, 80);
presentation.Save("Review.key");
#endregion
