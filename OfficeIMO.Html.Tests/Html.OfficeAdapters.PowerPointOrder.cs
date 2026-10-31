using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlOfficeAdaptersPowerPointOrder {
    [Fact]
    public void PowerPointHtml_RoundTripsShapesAtTheSlideOrigin() {
        using PowerPointDocument presentation = PowerPointDocument.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        slide.AddTextBox("Origin text",global::OfficeIMO.OfficeLength.Points(0),global::OfficeIMO.OfficeLength.Points(0),global::OfficeIMO.OfficeLength.Points(180),global::OfficeIMO.OfficeLength.Points(35));
        PowerPointTable table = slide.AddTable(1,1,global::OfficeIMO.OfficeLength.Points(0),global::OfficeIMO.OfficeLength.Points(0),global::OfficeIMO.OfficeLength.Points(220),global::OfficeIMO.OfficeLength.Points(60));
        table.GetCell(0, 0).Text = "Origin table";

        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides
        });
        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointDocumentResult();
        using PowerPointDocument imported = result.Value;
        PowerPointSlide importedSlide = Assert.Single(imported.Slides);

        Assert.Contains("data-officeimo-left=\"0\" data-officeimo-top=\"0\"", html, StringComparison.Ordinal);
        PowerPointTextBox importedText = Assert.Single(importedSlide.TextBoxes);
        PowerPointTable importedTable = Assert.Single(importedSlide.Tables);
        Assert.Equal(0D, importedText.Left.ToPoints(), 3);
        Assert.Equal(0D, importedText.Top.ToPoints(), 3);
        Assert.Equal(0D, importedTable.Left.ToPoints(), 3);
        Assert.Equal(0D, importedTable.Top.ToPoints(), 3);
        Assert.Empty(result.Report.Diagnostics);
    }

    [Fact]
    public void PowerPointHtml_RoundTripsUnifiedShapeReadingOrderAndGeometry() {
        using PowerPointDocument presentation = PowerPointDocument.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointTextBox firstText = slide.AddTextBox("First text",global::OfficeIMO.OfficeLength.Points(30),global::OfficeIMO.OfficeLength.Points(40),global::OfficeIMO.OfficeLength.Points(180),global::OfficeIMO.OfficeLength.Points(35));
        using (var image = new MemoryStream(Convert.FromBase64String(
                   "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNg+P//HwAF/gL9HjcXBgAAAABJRU5ErkJggg=="))) {
            slide.AddImage(image, OfficeIMO.Drawing.OfficeImageFormat.Png, OfficeIMO.OfficeLength.Points(60), OfficeIMO.OfficeLength.Points(90), OfficeIMO.OfficeLength.Points(70), OfficeIMO.OfficeLength.Points(50)).Name = "Middle picture";
        }

        PowerPointTable table = slide.AddTable(1,1,global::OfficeIMO.OfficeLength.Points(90),global::OfficeIMO.OfficeLength.Points(150),global::OfficeIMO.OfficeLength.Points(220),global::OfficeIMO.OfficeLength.Points(60));
        table.GetCell(0, 0).Text = "Ordered table";
        PowerPointTextBox lastText = slide.AddTextBox("Last text",global::OfficeIMO.OfficeLength.Points(120),global::OfficeIMO.OfficeLength.Points(230),global::OfficeIMO.OfficeLength.Points(190),global::OfficeIMO.OfficeLength.Points(40));
        table.MoveToReadingOrder(0);

        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides
        });
        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointDocumentResult();
        using PowerPointDocument imported = result.Value;
        PowerPointSlide importedSlide = Assert.Single(imported.Slides);
        PowerPointShape[] orderedShapes = importedSlide.Shapes.OrderBy(shape => shape.DrawingOrder).ToArray();

        Assert.Collection(
            orderedShapes,
            shape => Assert.IsType<PowerPointTable>(shape),
            shape => Assert.Same(firstText.GetType(), shape.GetType()),
            shape => Assert.IsType<PowerPointPicture>(shape),
            shape => Assert.Same(lastText.GetType(), shape.GetType()));
        PowerPointTable importedTable = Assert.IsType<PowerPointTable>(orderedShapes[0]);
        Assert.Equal(90D, importedTable.Left.ToPoints(), 3);
        Assert.Equal(150D, importedTable.Top.ToPoints(), 3);
        PowerPointTextBox importedLastText = Assert.IsType<PowerPointTextBox>(orderedShapes[3]);
        Assert.Equal(120D, importedLastText.Left.ToPoints(), 3);
        Assert.Equal(230D, importedLastText.Top.ToPoints(), 3);
        Assert.Empty(result.Report.Diagnostics);
    }

    [Fact]
    public void PowerPointHtml_UsesTheLastStandaloneNotesMarkerAfterManyFalseCandidates() {
        string falseCandidates = string.Join("\n", Enumerable.Repeat("prefix ### Notes suffix", 10_000));
        string html = "<section class='officeimo-slide'><pre class='officeimo-source-markdown'>"
            + falseCandidates + "\n### Notes\nBounded presenter notes</pre></section>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html).ToPowerPointDocumentResult();
        using PowerPointDocument presentation = result.Value;

        Assert.Equal("Bounded presenter notes", Assert.Single(presentation.Slides).Notes.Text);
    }
}
