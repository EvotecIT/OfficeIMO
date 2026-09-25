using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlOfficeAdaptersPowerPointGenericPaginationTests {
    [Fact]
    public void LongGenericArticleContinuesOnVisibleSlidesAndSurvivesReopen() {
        const string image = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNg+P//HwAF/gL9HjcXBgAAAABJRU5ErkJggg==";
        string html = "<article><h1>Field notes</h1>"
            + string.Concat(Enumerable.Range(1, 18).Select(index => "<p>Entry " + index + "</p>"))
            + "<p>Illustration <img src='" + image + "' alt='Field diagram'></p></article>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult(
            new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic, ImportEditableLayoutRegions = false });
        using PowerPointPresentation presentation = result.Value;

        Assert.True(result.Slides >= 3);
        Assert.Equal(18, presentation.Slides.Sum(slide =>
            slide.TextBoxes.Count(box => box.Text.StartsWith("Entry ", StringComparison.Ordinal))));
        Assert.Equal(1, presentation.Slides.Sum(slide => slide.Pictures.Count()));
        Assert.All(presentation.Slides, slide => Assert.All(slide.Shapes, shape => {
            Assert.True(shape.Bounds.Left >= 0 && shape.Bounds.Top >= 0);
            Assert.True(shape.Bounds.Right <= presentation.SlideSize.WidthEmus);
            Assert.True(shape.Bounds.Bottom <= presentation.SlideSize.HeightEmus);
        }));
        Assert.DoesNotContain(result.Report.Diagnostics, item => item.Code == HtmlConversionDiagnosticCodes.ContentOmitted);

        using var stream = new MemoryStream();
        presentation.Save(stream);
        using PowerPointPresentation reopened = PowerPointPresentation.Load(new MemoryStream(stream.ToArray()),
            new PowerPointLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
        Assert.Equal(presentation.Slides.Count, reopened.Slides.Count);
        Assert.Contains(reopened.Slides.SelectMany(slide => slide.TextBoxes), box => box.Text == "Entry 18");
        Assert.Equal(1, reopened.Slides.Sum(slide => slide.Pictures.Count()));
    }

    [Fact]
    public void ContinuationSlidesRespectTheSharedContainerLimit() {
        string html = "<article><h1>Field notes</h1>"
            + string.Concat(Enumerable.Range(1, 18).Select(index => "<p>Entry " + index + "</p>"))
            + "</article>";
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxSemanticContainers = 2;

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult(
            new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic, Limits = limits,
                ImportEditableLayoutRegions = false });
        using PowerPointPresentation presentation = result.Value;

        Assert.Equal(2, presentation.Slides.Count);
        Assert.Contains(result.Report.Diagnostics, item =>
            item.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded
            && item.LossKind == OfficeConversionLossKind.Omission);
    }

    [Fact]
    public void OversizedSingleCellTableRemainsVisibleWithStructureLossReported() {
        string cellText = string.Join(' ', Enumerable.Repeat("Gallery entry with a description and source link.", 24));
        string html = "<article><h1>Gallery</h1><table><tr><td><a href='https://example.org/gallery'>Gallery source</a> " + cellText
            + "</td></tr></table><p>After gallery</p></article>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult(
            new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic, ImportEditableLayoutRegions = false });
        using PowerPointPresentation presentation = result.Value;

        Assert.True(presentation.Slides.Count > 1);
        Assert.Empty(presentation.Slides.SelectMany(slide => slide.Tables));
        Assert.Contains(presentation.Slides.SelectMany(slide => slide.TextBoxes), box =>
            box.Text.Contains("Gallery entry with a description", StringComparison.Ordinal));
        Assert.Contains(presentation.Slides.SelectMany(slide => slide.TextBoxes), box => box.Text == "After gallery");
        Assert.Contains(result.Report.Diagnostics, item =>
            item.Code == HtmlConversionDiagnosticCodes.ContentApproximated
            && item.Detail?.Contains("projection=paginatedText", StringComparison.Ordinal) == true);
        Assert.Contains(result.Report.Diagnostics, item =>
            item.Code == HtmlConversionDiagnosticCodes.ContentOmitted
            && item.Detail?.Contains("hyperlinkRuns=", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void LayoutRegionKeepsItsSectionAfterEarlierSlideContinuation() {
        string first = "<section>" + string.Concat(Enumerable.Range(1, 12)
            .Select(index => "<p>First section entry " + index + "</p>")) + "</section>";
        const string second = "<section style='position:absolute;width:180px;height:50px'>"
            + "Second section region</section>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(first + second)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.Value;

        Assert.True(presentation.Slides.Count >= 3);
        Assert.Contains(presentation.Slides[^1].TextBoxes, box => box.Text == "Second section region");
        Assert.DoesNotContain(result.Report.Diagnostics, item =>
            item.Message.Contains("owning semantic slide was not created", StringComparison.Ordinal));
    }
}
