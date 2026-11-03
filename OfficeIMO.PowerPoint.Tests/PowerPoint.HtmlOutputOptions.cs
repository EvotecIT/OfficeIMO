using OfficeIMO.Drawing;
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
        var options = new PowerPointHtmlExportOptions { ExportProfile = profile };

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
        var options = new PowerPointHtmlExportOptions { ExportProfile = profile };

        string withHeadings = presentation.ToHtml(options);
        options.IncludeSlideHeadings = false;
        string withoutHeadings = presentation.ToHtml(options);

        Assert.Contains("<h2>Slide 1</h2>", withHeadings, StringComparison.Ordinal);
        Assert.Equal(withHeadings.Replace("<h2>Slide 1</h2>", string.Empty), withoutHeadings);
    }

    [Fact]
    public void OmittedHeadingIgnoresTitle() {
        using var presentation = CreateWithOffSlideShape();
        var options = new PowerPointHtmlExportOptions { Title = "Quarterly review", IncludeHeading = false };

        string html = presentation.ToHtml(options);

        Assert.DoesNotContain("<h1>", html, StringComparison.Ordinal);
        Assert.Contains("<title>Quarterly review</title>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void RenderDiagnosticsAreVisibleByDefaultAndCanBeLeftOutWhileStillReported() {
        using var presentation = CreateWithOffSlideShape();
        PowerPointHtmlExportOptions options = PowerPointHtmlExportOptions.CreateVisualReviewProfile();

        string visible = presentation.ToHtml(options, out OfficeConversionReport visibleReport);
        options.IncludeRenderDiagnostics = false;
        string omitted = presentation.ToHtml(options, out OfficeConversionReport omittedReport);

        Assert.Contains("<ul class=\"officeimo-snapshot-diagnostics\">", visible, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-diagnostic-code=\"invalid-shape-bounds\"", visible, StringComparison.Ordinal);
        Assert.DoesNotContain("officeimo-snapshot-diagnostics\"", omitted, StringComparison.Ordinal);
        Assert.DoesNotContain("invalid-shape-bounds", omitted, StringComparison.Ordinal);
        Assert.Contains(omittedReport.Diagnostics, _ => _.Code == "invalid-shape-bounds");
        Assert.Equal(visibleReport.Diagnostics.Select(_ => _.Code), omittedReport.Diagnostics.Select(_ => _.Code));
    }

    [Fact]
    public void CloneCopiesHeadingAndRenderDiagnosticChoices() {
        var options = new PowerPointHtmlExportOptions { IncludeHeading = false, IncludeSlideHeadings = false, IncludeRenderDiagnostics = false };

        PowerPointHtmlExportOptions clone = options.Clone();

        Assert.False(clone.IncludeHeading);
        Assert.False(clone.IncludeSlideHeadings);
        Assert.False(clone.IncludeRenderDiagnostics);
    }

    [Fact]
    public void VisualReviewSnapshotEmbedsSuppliedFonts() {
        using var presentation = CreateWithCalibriText();
        PowerPointHtmlExportOptions options = PowerPointHtmlExportOptions.CreateVisualReviewProfile();
        string installed = presentation.ToHtml(options);
        var fonts = new OfficeFontFaceCollection().Add("Calibri", ReadBaselineSans());
        options.Fonts = fonts;

        string supplied = presentation.ToHtml(options);

        Assert.Contains(SlideSvg(presentation, new PowerPointImageExportOptions { Fonts = fonts }), supplied, StringComparison.Ordinal);
        Assert.Equal(supplied, presentation.ToHtml(options.Clone()));
        Assert.NotEqual(installed, supplied);
    }

    [Fact]
    public void VisualReviewSnapshotKeepsPlatformFontsWithoutFonts() {
        using var presentation = CreateWithCalibriText();
        PowerPointHtmlExportOptions options = PowerPointHtmlExportOptions.CreateVisualReviewProfile();

        string html = presentation.ToHtml(options);

        Assert.Null(options.Fonts);
        Assert.Null(options.Clone().Fonts);
        Assert.Contains(SlideSvg(presentation, new PowerPointImageExportOptions()), html, StringComparison.Ordinal);
    }

    static string SlideSvg(PowerPointDocument presentation, PowerPointImageExportOptions options) =>
        OfficeDrawingSvgExporter.ToSvg(presentation.Slides[0].CreateVisualSnapshot(options).Drawing);

    static byte[] ReadBaselineSans() =>
        File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Fonts", "OfficeIMOBaselineSans-Regular.ttf"));

    static PowerPointDocument CreateWithCalibriText() {
        var presentation = PowerPointDocument.Create();
        presentation.SlideSize.SetSize(OfficeLength.Points(320), OfficeLength.Points(120));
        var box = presentation.AddSlide().AddTextBox("Pinned fonts", OfficeLength.Points(20), OfficeLength.Points(20), OfficeLength.Points(280), OfficeLength.Points(60));
        box.FontName = "Calibri";
        return presentation;
    }

    static PowerPointDocument CreateWithOffSlideShape() {
        var presentation = PowerPointDocument.Create();
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(160),OfficeIMO.OfficeLength.Points(100));
        presentation.AddSlide().AddRectangle(OfficeIMO.OfficeLength.Points(500), OfficeIMO.OfficeLength.Points(500), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(20), "Outside slide bounds");
        return presentation;
    }
}
