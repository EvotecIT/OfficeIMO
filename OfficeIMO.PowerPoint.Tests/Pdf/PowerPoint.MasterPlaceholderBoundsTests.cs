using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointMasterPlaceholderBoundsTests {
    [Fact]
    public void TitleWithoutSlideOrLayoutTransformUsesMasterBoundsAfterReopen() {
        using var bytes = new MemoryStream();
        using (var deck = PowerPointPresentation.Create()) {
            deck.SlideSize.SetSizePoints(400, 250);
            var slide = deck.AddSlide();
            var title = slide.AddTextBoxPoints("MASTER INHERITED TITLE", 10, 10, 200, 50);
            title.PlaceholderType = PowerPointPlaceholderType.Title;
            title.PlaceholderIndex = 0;
            title.FontName = "Helvetica";
            title.FontSize = 18;
            Assert.IsType<Shape>(title.Element).ShapeProperties!.Transform2D = null;
            var layout = slide.SlidePart.SlideLayoutPart!.SlideLayout.CommonSlideData!.ShapeTree!;
            foreach (var shape in layout.Elements<Shape>().ToArray()) shape.Remove();
            var inherited = (Shape)title.Element.CloneNode(true);
            inherited.ShapeProperties!.Transform2D = null;
            layout.Append(inherited);
            var master = slide.SlidePart.SlideLayoutPart.SlideMasterPart!.SlideMaster.CommonSlideData!.ShapeTree!;
            foreach (var shape in master.Elements<Shape>().Where(shape => shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == PlaceholderValues.Title).ToArray()) shape.Remove();
            var masterShape = (Shape)inherited.CloneNode(true);
            masterShape.ShapeProperties!.Transform2D = new A.Transform2D(new A.Offset { X = 254000, Y = 381000 }, new A.Extents { Cx = 3810000, Cy = 635000 });
            master.Append(masterShape);
            deck.Save(bytes);
        }
        bytes.Position = 0;
        using var reopened = PowerPointPresentation.Load(bytes);
        var conversion = reopened.ToPdfDocumentResult();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(conversion.ToBytes());
        Assert.Contains("MASTER INHERITED TITLE", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain(conversion.Warnings, warning => warning.Code == "invalid-shape-bounds");
        Assert.InRange(pdf.GetPage(1).Letters.Min(letter => letter.StartBaseLine.X), 20D, 30D);
    }
}
