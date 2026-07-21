using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlFidelityScoreV2 {
    [Fact]
    public void FidelityV2_ScoresStylesResourcesAnnotationsFormulasChartsAndGeometry() {
        const string source = """
            <main>
              <p style="color:#123456"><strong>Styled</strong></p>
              <img src="https://example.test/a.png" alt="Chart">
              <ins data-officeimo-annotation="review">Added</ins>
              <section class="officeimo-formulas"><ul><li data-officeimo-cell="A1"><code>=SUM(B1:B2)</code></li></ul></section>
              <section class="officeimo-charts"><ul><li data-officeimo-chart-type="Bar" data-officeimo-left="10" data-officeimo-top="20" data-officeimo-width="300" data-officeimo-height="180">Revenue</li></ul></section>
            </main>
            """;
        const string degraded = """
            <main>
              <p style="color:#654321">Styled</p>
              <img src="https://example.test/b.png" alt="Different">
              <del data-officeimo-annotation="review">Removed</del>
              <section class="officeimo-formulas"><ul><li data-officeimo-cell="A1"><code>=B1</code></li></ul></section>
              <section class="officeimo-charts"><ul><li data-officeimo-chart-type="Line" data-officeimo-left="25" data-officeimo-top="40" data-officeimo-width="500" data-officeimo-height="200">Costs</li></ul></section>
            </main>
            """;

        HtmlRoundTripScore exact = HtmlRoundTripScorer.Compare(source, source);
        HtmlRoundTripScore lossy = HtmlRoundTripScorer.Compare(source, degraded);

        Assert.Equal(HtmlRoundTripScore.CurrentSchemaVersion, exact.SchemaVersion);
        foreach (string dimension in new[] { "structure", "text", "styles", "resources", "annotations", "formulas", "charts", "geometry" }) {
            Assert.Equal(1D, exact.Dimensions[dimension], 12);
        }
        Assert.InRange(lossy.Dimensions["styles"], 0D, 0.99D);
        Assert.InRange(lossy.Dimensions["resources"], 0D, 0.99D);
        Assert.InRange(lossy.Dimensions["annotations"], 0D, 0.99D);
        Assert.InRange(lossy.Dimensions["formulas"], 0D, 0.99D);
        Assert.InRange(lossy.Dimensions["charts"], 0D, 0.99D);
        Assert.InRange(lossy.Dimensions["geometry"], 0D, 0.99D);
        Assert.True(lossy.Score < exact.Score);
    }

    [Fact]
    public void FidelityV2_UsesHtmlExportedAfterActualDocxReload() {
        const string html = "<h1>Reload proof</h1><p>Plain <strong>bold</strong> <a href='https://example.com'>link</a></p><ul><li>One</li><li>Two</li></ul>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        using WordDocument document = source.ToWordDocumentResult().RequireValue();
        string initialExport = document.ToHtml();
        using MemoryStream artifact = document.ToStream();
        using WordDocument reloaded = WordDocument.Load(
            new MemoryStream(artifact.ToArray()),
            new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
        string reloadedExport = reloaded.ToHtml();

        HtmlRoundTripScore score = HtmlRoundTripScorer.Compare(
            source,
            HtmlConversionDocument.Parse(initialExport),
            HtmlArtifactReloadEvidence.Succeeded("DOCX", reloadedExport));

        Assert.True(score.ArtifactReloadVerified);
        Assert.Equal("DOCX", score.ArtifactKind);
        Assert.Equal(score.Metrics["artifact-reload"], score.Dimensions["artifact-reload"], 12);
        Assert.InRange(score.Dimensions["artifact-reload"], 0.50D, 1D);
    }

    [Fact]
    public void FidelityV2_ReportsFailedReloadEvidenceWithoutClaimingVerification() {
        HtmlRoundTripScore score = HtmlRoundTripScorer.Compare(
            "<p>Source</p>",
            "<p>Source</p>",
            HtmlArtifactReloadEvidence.Failed("PPTX", "package validation failed"));

        Assert.False(score.ArtifactReloadVerified);
        Assert.Equal(0D, score.Dimensions["artifact-reload"]);
        Assert.Equal("PPTX", score.ArtifactKind);
    }

    [Fact]
    public void FidelityV2_UsesHtmlExportedAfterActualXlsxReload() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<table><caption>Data</caption><tr><th>Name</th><th>Value</th></tr><tr><td>Total</td><td>42</td></tr></table>");
        using ExcelDocument workbook = source.ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic }).RequireValue();
        string initialExport = workbook.ToHtml();
        using MemoryStream artifact = workbook.ToStream();
        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(artifact.ToArray()),
            new ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });

        HtmlRoundTripScore score = HtmlRoundTripScorer.Compare(source,
            HtmlConversionDocument.Parse(initialExport),
            HtmlArtifactReloadEvidence.Succeeded("XLSX", reloaded.ToHtml()));

        Assert.True(score.ArtifactReloadVerified);
        Assert.Equal("XLSX", score.ArtifactKind);
        Assert.InRange(score.Dimensions["artifact-reload"], 0.01D, 1D);
    }

    [Fact]
    public void FidelityV2_UsesHtmlExportedAfterActualPptxReload() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<section><h1>Reload</h1><p>Plain <strong>bold</strong></p></section>");
        using PowerPointPresentation presentation = source.ToPowerPointPresentationResult(
            new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic }).RequireValue();
        string initialExport = presentation.ToHtml();
        using MemoryStream artifact = presentation.ToStream();
        using PowerPointPresentation reloaded = PowerPointPresentation.Load(
            new MemoryStream(artifact.ToArray()),
            new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });

        HtmlRoundTripScore score = HtmlRoundTripScorer.Compare(source,
            HtmlConversionDocument.Parse(initialExport),
            HtmlArtifactReloadEvidence.Succeeded("PPTX", reloaded.ToHtml()));

        Assert.True(score.ArtifactReloadVerified);
        Assert.Equal("PPTX", score.ArtifactKind);
        Assert.InRange(score.Dimensions["artifact-reload"], 0.01D, 1D);
    }
}
