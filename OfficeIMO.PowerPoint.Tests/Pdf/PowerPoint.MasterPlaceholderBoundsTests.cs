using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointMasterPlaceholderBoundsTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LayoutPlaceholderPromptsAreNotContentAndConcreteSlidePlaceholdersStillRender(bool concrete) {
        using var deck = PowerPointDocument.Create();
        deck.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(400),OfficeIMO.OfficeLength.Points(250));
        var slide = deck.AddSlide();
        var tree = slide.SlidePart.SlideLayoutPart!.SlideLayout.CommonSlideData!.ShapeTree!;
        PlaceholderValues?[] types = { PlaceholderValues.DateAndTime, PlaceholderValues.Footer, PlaceholderValues.SlideNumber, null };
        for (int index = 0; index < types.Length; index++) {
            var prompt = slide.AddTextBox("TEMPLATE PROMPT " + index,global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(20 + index * 40),global::OfficeIMO.OfficeLength.Points(300),global::OfficeIMO.OfficeLength.Points(30));
            Assert.IsType<Shape>(prompt.Element).NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
                .PlaceholderShape = types[index].HasValue
                    ? new PlaceholderShape { Type = types[index]!.Value, Index = (uint)(10 + index) }
                    : new PlaceholderShape();
            prompt.Element.Remove();
            tree.Append(prompt.Element);
            if (concrete) {
                var actual = slide.AddTextBox("ACTUAL FIELD " + index,global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(20 + index * 40),global::OfficeIMO.OfficeLength.Points(300),global::OfficeIMO.OfficeLength.Points(30));
                Assert.IsType<Shape>(actual.Element).NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
                    .PlaceholderShape = types[index].HasValue
                        ? new PlaceholderShape { Type = types[index]!.Value, Index = (uint)(10 + index) }
                        : new PlaceholderShape();
            }
        }
        var artwork = slide.AddTextBox("LAYOUT ARTWORK",global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(180),global::OfficeIMO.OfficeLength.Points(300),global::OfficeIMO.OfficeLength.Points(30));
        artwork.PlaceholderType = null;
        artwork.Element.Remove();
        tree.Append(artwork.Element);

        using var saved = new MemoryStream();
        deck.Save(saved);
        saved.Position = 0;
        using var reopened = PowerPointDocument.Load(saved);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(reopened.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        string text = pdf.GetPage(1).Text;
        Assert.DoesNotContain("TEMPLATE PROMPT", text, StringComparison.Ordinal);
        Assert.Contains("LAYOUT ARTWORK", text, StringComparison.Ordinal);
        for (int index = 0; index < types.Length; index++)
            Assert.Equal(concrete, text.Contains("ACTUAL FIELD " + index, StringComparison.Ordinal));
    }

    [Fact]
    public void TitleWithoutSlideOrLayoutTransformUsesMasterBoundsAfterReopen() {
        using var bytes = new MemoryStream();
        using (var deck = PowerPointDocument.Create()) {
            deck.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(400),OfficeIMO.OfficeLength.Points(250));
            var slide = deck.AddSlide();
            var title = slide.AddTextBox("MASTER INHERITED TITLE",global::OfficeIMO.OfficeLength.Points(10),global::OfficeIMO.OfficeLength.Points(10),global::OfficeIMO.OfficeLength.Points(200),global::OfficeIMO.OfficeLength.Points(50));
            title.PlaceholderType = PowerPointPlaceholderType.Title;
            title.PlaceholderIndex = 0;
            title.FontName = "Helvetica";
            title.FontSize = global::OfficeIMO.OfficeLength.Points(18);
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
        using var reopened = PowerPointDocument.Load(bytes);
        var conversion = reopened.ConvertToPdfWithEvidence();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(conversion.ToBytes());
        Assert.Contains("MASTER INHERITED TITLE", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain(conversion.Warnings, warning => warning.Code == "invalid-shape-bounds");
        Assert.InRange(pdf.GetPage(1).Letters.Min(letter => letter.StartBaseLine.X), 20D, 30D);
    }
}
