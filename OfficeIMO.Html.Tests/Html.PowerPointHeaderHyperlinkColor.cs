using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlPowerPointHeaderHyperlinkColorTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void NativeAndRepeatedHeadersKeepTheirTextColorWithoutChangingTargets(bool paginate, bool authoredColor) {
        string rows = string.Concat(Enumerable.Range(1, paginate ? 30 : 1).Select(index =>
            "<tr><td>Row " + index + "</td><td>" + (paginate
                ? string.Join(" ", Enumerable.Repeat("A native editable table value remains in source order.", 4))
                : "Value") + "</td></tr>"));
        string html = "<table><thead><tr><th>Label</th><th"
            + (authoredColor ? " style='background-color:#123456;color:#ffdd00'" : "")
            + "><a href='https://example.test/source'>Source</a></th></tr></thead><tbody>" + rows + "</tbody></table>";
        using PowerPointPresentation presentation = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions {
                Mode = HtmlImportMode.Generic, ImportEditableLayoutRegions = false
            }).RequireValue();
        var tables = presentation.Slides.SelectMany(slide => slide.Tables).ToArray();
        if (paginate) Assert.True(tables.Length > 1);
        Assert.All(tables, table => {
            var run = Assert.Single(table.GetCell(0, 1).Runs);
            Assert.Equal("Source", run.Text);
            Assert.Equal("https://example.test/source", run.Hyperlink!.AbsoluteUri);
            Assert.True(run.HyperlinkUsesTextColor);
            if (authoredColor) Assert.Equal("FFDD00", run.Color);
        });
        Assert.Empty(new OpenXmlValidator().Validate(presentation.OpenXmlDocument));
        using var stream = new MemoryStream();
        presentation.Save(stream);
        stream.Position = 0;
        using var reopened = PowerPointPresentation.Load(stream);
        Assert.All(reopened.Slides.SelectMany(slide => slide.Tables), table =>
            Assert.True(table.GetCell(0, 1).Runs.Single().HyperlinkUsesTextColor));
        Assert.Equal(Enumerable.Range(1, paginate ? 30 : 1).Select(index => "Row " + index),
            reopened.Slides.SelectMany(slide => slide.Tables).SelectMany(table =>
                Enumerable.Range(1, table.Rows - 1).Select(row => table.GetCell(row, 0).Text)));
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SemanticRoundTripPreservesExplicitHyperlinkColorPolicy(bool textColor) {
        using var source = PowerPointPresentation.Create();
        var run = source.AddSlide().AddTextBox("Source link").Paragraphs.Single().Runs.Single();
        run.Color = "FFDD00";
        run.SetHyperlink("https://example.test/source");
        run.HyperlinkUsesTextColor = textColor;
        using PowerPointPresentation imported = HtmlConversionDocument.Parse(source.ToHtml())
            .ToPowerPointPresentationResult().RequireValue();
        var actual = imported.Slides.Single().TextBoxes.Single().Paragraphs.Single().Runs.Single();
        Assert.Equal(textColor, actual.HyperlinkUsesTextColor);
        Assert.Equal("FFDD00", actual.Color);
        Assert.Equal("https://example.test/source", actual.Hyperlink!.AbsoluteUri);
        Assert.Equal("Source link", actual.Text);
        Assert.Empty(new OpenXmlValidator().Validate(imported.OpenXmlDocument));
    }

}
