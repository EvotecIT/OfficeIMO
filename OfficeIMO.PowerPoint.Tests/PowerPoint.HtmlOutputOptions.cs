using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointHtmlOutputOptionsTests {
    [Theory]
    [InlineData(PowerPointHtmlExportProfile.SemanticSlides, "PowerPoint Presentation")]
    [InlineData(PowerPointHtmlExportProfile.VisualReview, "PowerPoint Visual Review")]
    public void HeadingIsWrittenByDefaultAndCanBeOmitted(PowerPointHtmlExportProfile profile, string fallback) {
        using var presentation = CreateWithOffSlideShape();
        var options = new PowerPointHtmlSaveOptions { ExportProfile = profile };

        string withHeading = presentation.ToHtml(options);
        options.IncludeHeading = false;
        string withoutHeading = presentation.ToHtml(options);

        Assert.Contains("<h1>" + fallback + "</h1>", withHeading, StringComparison.Ordinal);
        Assert.DoesNotContain("<h1>", withoutHeading, StringComparison.Ordinal);
        // The heading is the only difference; the document title is still set.
        Assert.Contains("<title>" + fallback + "</title>", withoutHeading, StringComparison.Ordinal);
        Assert.Equal(withHeading.Replace("<h1>" + fallback + "</h1>", string.Empty), withoutHeading);
    }

    [Theory]
    [InlineData(PowerPointHtmlExportProfile.SemanticSlides)]
    [InlineData(PowerPointHtmlExportProfile.VisualReview)]
    public void SlideHeadingsAreWrittenByDefaultAndCanBeOmitted(PowerPointHtmlExportProfile profile) {
        using var presentation = CreateWithOffSlideShape();
        var options = new PowerPointHtmlSaveOptions { ExportProfile = profile };

        string withHeadings = presentation.ToHtml(options);
        options.IncludeSlideHeadings = false;
        string withoutHeadings = presentation.ToHtml(options);

        Assert.Contains("<h2>Slide 1</h2>", withHeadings, StringComparison.Ordinal);
        Assert.Equal(withHeadings.Replace("<h2>Slide 1</h2>", string.Empty), withoutHeadings);
    }

    [Fact]
    public void OmittedHeadingIgnoresTitle() {
        using var presentation = CreateWithOffSlideShape();
        var options = new PowerPointHtmlSaveOptions { Title = "Quarterly review", IncludeHeading = false };

        string html = presentation.ToHtml(options);

        Assert.DoesNotContain("<h1>", html, StringComparison.Ordinal);
        Assert.Contains("<title>Quarterly review</title>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void RenderDiagnosticsAreVisibleByDefaultAndCanBeLeftOutWhileStillReported() {
        using var presentation = CreateWithOffSlideShape();
        PowerPointHtmlSaveOptions options = PowerPointHtmlSaveOptions.CreateVisualReviewProfile();

        PowerPointToHtmlResult visible = presentation.ToHtmlResult(options);
        options.IncludeRenderDiagnostics = false;
        PowerPointToHtmlResult omitted = presentation.ToHtmlResult(options);

        Assert.Contains("<ul class=\"officeimo-snapshot-diagnostics\">", visible.Value, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-diagnostic-code=\"invalid-shape-bounds\"", visible.Value, StringComparison.Ordinal);
        Assert.DoesNotContain("officeimo-snapshot-diagnostics\"", omitted.Value, StringComparison.Ordinal);
        Assert.DoesNotContain("invalid-shape-bounds", omitted.Value, StringComparison.Ordinal);
        Assert.Contains(omitted.ImageDiagnostics, _ => _.Code == "invalid-shape-bounds");
        Assert.Equal(visible.ImageDiagnostics.Select(_ => _.Code), omitted.ImageDiagnostics.Select(_ => _.Code));
    }

    [Fact]
    public void CloneCopiesHeadingAndRenderDiagnosticChoices() {
        var options = new PowerPointHtmlSaveOptions { IncludeHeading = false, IncludeSlideHeadings = false, IncludeRenderDiagnostics = false };

        PowerPointHtmlSaveOptions clone = options.Clone();

        Assert.False(clone.IncludeHeading);
        Assert.False(clone.IncludeSlideHeadings);
        Assert.False(clone.IncludeRenderDiagnostics);
    }

    static PowerPointDocument CreateWithOffSlideShape() {
        var presentation = PowerPointDocument.Create();
        presentation.SlideSize.SetSizePoints(160, 100);
        presentation.AddSlide().AddRectanglePoints(500, 500, 20, 20, "Outside slide bounds");
        return presentation;
    }
}
