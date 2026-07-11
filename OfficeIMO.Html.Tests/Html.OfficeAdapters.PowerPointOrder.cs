using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlOfficeAdaptersPowerPointOrder {
    [Fact]
    public void PowerPointHtml_RoundTripsUnifiedShapeReadingOrderAndGeometry() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.Slides[0];
        PowerPointTextBox firstText = slide.AddTextBoxPoints("First text", 30, 40, 180, 35);
        using (var image = new MemoryStream(Convert.FromBase64String(
                   "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAEAQH/69DjmQAAAABJRU5ErkJggg=="))) {
            slide.AddPicturePoints(image, ImagePartType.Png, 60, 90, 70, 50).Name = "Middle picture";
        }

        PowerPointTable table = slide.AddTablePoints(1, 1, 90, 150, 220, 60);
        table.GetCell(0, 0).Text = "Ordered table";
        PowerPointTextBox lastText = slide.AddTextBoxPoints("Last text", 120, 230, 190, 40);
        table.MoveToReadingOrder(0);

        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides
        });
        HtmlToPowerPointResult result = html.ToPowerPointPresentationResult();
        using PowerPointPresentation imported = result.Presentation;
        PowerPointSlide importedSlide = Assert.Single(imported.Slides);
        PowerPointShape[] orderedShapes = importedSlide.Shapes.OrderBy(shape => shape.DrawingOrder).ToArray();

        Assert.Collection(
            orderedShapes,
            shape => Assert.IsType<PowerPointTable>(shape),
            shape => Assert.Same(firstText.GetType(), shape.GetType()),
            shape => Assert.IsType<PowerPointPicture>(shape),
            shape => Assert.Same(lastText.GetType(), shape.GetType()));
        PowerPointTable importedTable = Assert.IsType<PowerPointTable>(orderedShapes[0]);
        Assert.Equal(90D, importedTable.LeftPoints, 3);
        Assert.Equal(150D, importedTable.TopPoints, 3);
        PowerPointTextBox importedLastText = Assert.IsType<PowerPointTextBox>(orderedShapes[3]);
        Assert.Equal(120D, importedLastText.LeftPoints, 3);
        Assert.Equal(230D, importedLastText.TopPoints, 3);
        Assert.Empty(result.Diagnostics);
    }
}
