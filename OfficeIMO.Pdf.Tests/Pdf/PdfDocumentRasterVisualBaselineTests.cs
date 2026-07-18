using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentRasterVisualBaselineTests {
    // Poppler/font rasterization can move a few glyph or image-edge pixels without changing layout.
    private const int DefaultAllowedRasterNoisePixels = 32;

    [Fact]
    public void ProfessionalReport_MatchesPopplerRasterBaseline() {
        AssertScenarioRasterBaseline("professional-report", CreateProfessionalReport);
    }

    [Fact]
    public void LineItemsTwoPage_MatchesPopplerRasterBaseline() {
        AssertScenarioRasterBaseline("line-items-two-page", CreateLineItemsTwoPage, pageCount: 2);
    }

    [Fact]
    public void HeadersFooters_MatchesPopplerRasterBaseline() {
        AssertScenarioRasterBaseline("headers-footers", CreateHeadersFooters, pageCount: 2);
    }

    [Fact]
    public void FlowDsl_MatchesPopplerRasterBaseline() {
        AssertScenarioRasterBaseline("flow-dsl", CreateFlowDsl, pageCount: 3);
    }

    [Fact]
    public void NativeWordReport_MatchesPopplerRasterBaseline() {
        AssertScenarioRasterBaseline("native-word-report", CreateNativeWordReport);
    }

    [Fact]
    public void NativeWordDailyLayout_MatchesPopplerRasterBaseline() {
        AssertScenarioRasterBaseline("native-word-daily-layout", CreateNativeWordDailyLayout);
    }

    [Fact]
    public void NativeWordTableCellPictureControl_MatchesPopplerRasterBaseline() {
        AssertScenarioRasterBaseline("native-word-table-cell-picture-control", CreateNativeWordTableCellPictureControl);
    }

    [Fact]
    public void NativeExcelDailyWorkbook_MatchesPopplerRasterBaseline() {
        AssertScenarioRasterBaseline("native-excel-daily-workbook", CreateNativeExcelDailyWorkbook, pageCount: 2);
    }

    [Fact]
    public void MarkdownTechnicalDocument_MatchesPopplerRasterBaseline() {
        AssertScenarioRasterBaseline("markdown-technical-document", CreateMarkdownTechnicalDocument);
    }

    [Theory]
    [InlineData(OfficeVisualThemeKind.Plain, "markdown-theme-gallery-plain")]
    [InlineData(OfficeVisualThemeKind.WordLike, "markdown-theme-gallery-word-like")]
    [InlineData(OfficeVisualThemeKind.TechnicalDocument, "markdown-theme-gallery-technical-document")]
    [InlineData(OfficeVisualThemeKind.GitHubLike, "markdown-theme-gallery-github-like")]
    [InlineData(OfficeVisualThemeKind.Compact, "markdown-theme-gallery-compact")]
    [InlineData(OfficeVisualThemeKind.Report, "markdown-theme-gallery-report")]
    public void MarkdownThemeGallery_MatchesPopplerRasterBaseline(OfficeVisualThemeKind themeKind, string scenarioName) {
        AssertScenarioRasterBaseline(scenarioName, () => CreateMarkdownThemeGallery(themeKind));
    }

    [Fact]
    public void NativeWordReport_ExposesBodyCheckBoxesAsAcroFormFields() {
        byte[] bytes = CreateNativeWordReport();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        Assert.Collection(
            info.FormFields.OrderBy(field => field.Name, StringComparer.Ordinal),
            approved => {
                Assert.Equal("NativeApproved", approved.Name);
                Assert.Equal(PdfFormFieldKind.Button, approved.Kind);
                Assert.True(approved.IsCheckBox);
                Assert.Equal("Yes", approved.Value);
            },
            deferred => {
                Assert.Equal("NativeDeferred", deferred.Name);
                Assert.Equal(PdfFormFieldKind.Button, deferred.Kind);
                Assert.True(deferred.IsCheckBox);
                Assert.Equal("Off", deferred.Value);
            },
            tableApproved => {
                Assert.Equal("NativeTableApproved", tableApproved.Name);
                Assert.Equal(PdfFormFieldKind.Button, tableApproved.Kind);
                Assert.True(tableApproved.IsCheckBox);
                Assert.Equal("Yes", tableApproved.Value);
            });
    }

    [Fact]
    public void NativeWordReport_ExposesTableListTocAndOutlineSignals() {
        byte[] bytes = CreateNativeWordReport();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        Assert.Equal("Native Word PDF Visual Gate", Assert.Single(info.Outlines).Title);
        Assert.Contains(info.Outlines[0].Children, outline => outline.Title == "Native proof areas" && outline.Level == 2);
        Assert.Contains(info.Outlines[0].Children, outline => outline.Title == "Native evidence table" && outline.Level == 2);

        PdfLogicalDocument logical = PdfLogicalDocument.Load(bytes, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        Assert.Contains(logical.Headings, heading => heading.Text == "Native Word PDF Visual Gate");
        Assert.Contains(logical.Headings, heading => heading.Text == "Native proof areas");
        Assert.Contains(logical.Headings, heading => heading.Text == "Native evidence table");

        var tocLinks = logical.GetLinksByDestinationName("officeimo-heading-native-word-pdf-visual-gate").ToList();
        Assert.NotEmpty(tocLinks);
        Assert.All(tocLinks, link => Assert.Equal("Table of contents: Native Word PDF Visual Gate", link.Contents));

        var listItems = PdfTextExtractor.ExtractListItemsByPage(bytes)
            .SelectMany(page => page.ListItems)
            .ToList();
        Assert.Contains(listItems, item => item.Text == "Native list mapping keeps markers and text aligned.");
        Assert.Contains(listItems, item => item.Marker == "1" && item.Text == "Generated TOC appears before content.");

        using UglyToad.PdfPig.PdfDocument pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("Table of Contents", text);
        Assert.Contains("AreaNative statusEvidence", text);
        Assert.Contains("TablesPartialstyle and borders", text);
    }

    [Fact]
    public void NativeWordDailyLayout_ExposesColumnsLinksTocAndLayoutSignals() {
        byte[] bytes = CreateNativeWordDailyLayout();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        Assert.Equal(1, info.PageCount);
        Assert.Contains(info.Outlines, outline => outline.Title == "Daily Word Layout Gate");
        Assert.Contains(info.LinkUris, uri => uri == "https://evotec.xyz/native-daily-layout");
        Assert.Contains(info.LinkUris, uri => uri == "https://officeimo.net/");
        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.918 0.957 1 rg", rawPdf, StringComparison.Ordinal);

        PdfLogicalDocument logical = PdfLogicalDocument.Load(bytes, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        Assert.NotEmpty(logical.GetLinksByDestinationName("officeimo-heading-daily-word-layout-gate"));

        using UglyToad.PdfPig.PdfDocument pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var page = pdf.GetPage(1);
        string pageText = page.Text;
        Assert.Contains("Table of Contents", pageText);
        Assert.Contains("Column narrative", pageText);
        Assert.Contains("Column evidence", pageText);
        Assert.Contains("Inline column break", pageText);
        Assert.Contains("separator", pageText);

        var words = page.GetWords().ToList();
        double narrativeX = words.First(word => word.Text == "narrative").BoundingBox.Left;
        double evidenceX = words.First(word => word.Text == "evidence").BoundingBox.Left;
        Assert.True(evidenceX > narrativeX + 250D, $"Expected evidence heading to render in the second Word section column. Narrative x: {narrativeX:0.##}, evidence x: {evidenceX:0.##}.");

        double rightMostTocLeaderDot = page.Letters
            .Where(letter => letter.Value == "." &&
                letter.StartBaseLine.Y > 635D &&
                letter.StartBaseLine.X < evidenceX)
            .Select(letter => letter.EndBaseLine.X)
            .DefaultIfEmpty(0D)
            .Max();
        Assert.True(rightMostTocLeaderDot < evidenceX - 12D, $"Expected table-of-contents dot leaders to stay inside the first Word section column. Leader right edge: {rightMostTocLeaderDot:0.##}, evidence x: {evidenceX:0.##}.");
    }

    [Fact]
    public void NativeExcelDailyWorkbook_ExposesSheetsLinksImagesAndLayoutSignals() {
        byte[] bytes = CreateNativeExcelDailyWorkbook();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        Assert.Equal(2, info.PageCount);
        Assert.Contains(info.Outlines, outline => outline.Title == "Summary");
        Assert.Contains(info.Outlines, outline => outline.Title == "Details");
        Assert.Contains(info.LinkUris, uri => uri == "https://officeimo.net/excel-pdf");

        PdfLogicalDocument logical = PdfLogicalDocument.Load(bytes, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        Assert.Contains(logical.NamedDestinations, destination => destination.Name.Contains("summary", StringComparison.OrdinalIgnoreCase));
        Assert.Single(logical.NamedDestinations, destination => string.Equals(destination.Name, "excel-sheet-2-details", StringComparison.Ordinal));
        PdfNamedDestination detailsCellDestination = Assert.Single(logical.NamedDestinations, destination => string.Equals(destination.Name, "excel-sheet-2-details-a1", StringComparison.Ordinal));
        PdfLogicalLinkAnnotation detailsLink = Assert.Single(logical.GetLinksByDestinationName(detailsCellDestination.Name));
        Assert.Equal("Open Details", detailsLink.Contents);

        var images = PdfImageExtractor.ExtractImages(bytes);
        string rawPdf = Encoding.ASCII.GetString(bytes);
        int imageDraws = Regex.Matches(rawPdf, @"/Im\d+\s+Do").Count;
        Assert.True(images.Count >= 1 && imageDraws >= 2, "Expected worksheet body and header image placements to survive Excel-to-PDF export.");

        using UglyToad.PdfPig.PdfDocument pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        UglyToad.PdfPig.Content.Page summaryPage = pdf.GetPage(1);
        string summaryText = summaryPage.Text;
        Assert.Contains("Daily Excel PDF Gate", summaryText);
        Assert.Contains("Revenue", summaryText);
        Assert.Contains("Revenue Chart", summaryText);
        Assert.Contains("Actual", summaryText);
        Assert.Contains("Target", summaryText);
        Assert.Contains("$12,345.60", summaryText);
        Assert.Contains("25.7%", summaryText);
        Assert.Contains("Open Details", summaryText);
        Assert.DoesNotContain("HiddenRowValue", summaryText);
        Assert.DoesNotContain("HiddenColumnValue", summaryText);

        double metricX = FindWordStartX(summaryPage, "Metric");
        double statusX = FindWordStartX(summaryPage, "Status");
        Assert.True(statusX > metricX + 160D, $"Expected explicit worksheet column widths to make the status column visibly farther right. Metric x: {metricX:0.##}, Status x: {statusX:0.##}.");

        string detailsText = pdf.GetPage(2).Text;
        Assert.Contains("Details", detailsText);
        Assert.Contains("Details Target", detailsText);
    }

    [Theory]
    [InlineData("hello-world")]
    [InlineData("core-layout")]
    [InlineData("style-cheatsheet")]
    [InlineData("links-rules")]
    [InlineData("lists-tables")]
    [InlineData("table-style-gallery")]
    [InlineData("default-styles")]
    [InlineData("styled-runs")]
    [InlineData("tabs-leaders")]
    [InlineData("drawing-gallery")]
    [InlineData("watermark")]
    [InlineData("image-watermark")]
    [InlineData("page-border")]
    [InlineData("background-image")]
    [InlineData("background-shapes")]
    [InlineData("row-columns")]
    [InlineData("showcase-dashboard")]
    public void CorePdfScenarios_MatchPopplerRasterBaseline(string scenarioName) {
        AssertScenarioRasterBaseline(scenarioName, () => CreateCoreScenario(scenarioName));
    }

    private static void AssertScenarioRasterBaseline(string scenarioName, Func<byte[]> createPdf, int pageCount = 1) {
        if (pageCount <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pageCount), pageCount, "Raster baseline page count must be positive.");
        }

        byte[] pdfBytes = createPdf();
        int actualPageCount = PdfInspector.Inspect(pdfBytes).PageCount;
        if (actualPageCount != pageCount) {
            throw new Xunit.Sdk.XunitException(
                "PDF raster baseline scenario '" + scenarioName + "' produced " +
                actualPageCount.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " page(s), but the approved baseline expects " +
                pageCount.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " page(s). Update the expected page count and baselines deliberately if this page-flow change is intended.");
        }

        WriteReviewPdfArtifact(scenarioName, pdfBytes);
        if (SkipRasterAssertions()) {
            return;
        }

        if (!TryFindPdftoppm(out string rasterizerPath)) {
            if (IsStrictRasterBaselineRequired()) {
                throw new InvalidOperationException("PDF raster baseline tests require Poppler pdftoppm. Install Poppler or set OFFICEIMO_PDF_RASTERIZER to pdftoppm.exe.");
            }

            return;
        }

        if (!CanAssertRasterBaseline(rasterizerPath)) {
            return;
        }

        string workDir = Path.Combine(Path.GetTempPath(), "OfficeIMO.PdfRaster", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workDir);
        string pdfPath = Path.Combine(workDir, scenarioName + ".pdf");

        try {
            File.WriteAllBytes(pdfPath, pdfBytes);
            for (int pageNumber = 1; pageNumber <= pageCount; pageNumber++) {
                string pageText = pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
                string outputPrefix = Path.Combine(workDir, scenarioName + "-page" + pageText);
                string actualPng = outputPrefix + ".png";

                RunPdftoppm(rasterizerPath, pdfPath, outputPrefix, workDir, pageNumber);

                if (!File.Exists(actualPng)) {
                    throw new FileNotFoundException("Poppler did not produce the expected PNG page snapshot.", actualPng);
                }

                AssertRasterBaseline("officeimo-pdf-" + scenarioName + ".page" + pageText + ".poppler.png", actualPng);
            }
        } finally {
            TryDeleteDirectory(workDir);
        }
    }

    [Fact]
    public void RasterBaseline_RejectsUnexpectedGeneratedPageCount() {
        var exception = Assert.Throws<Xunit.Sdk.XunitException>(() =>
            AssertScenarioRasterBaseline("page-count-mismatch", () =>
                PdfDocument.Create()
                    .Paragraph(p => p.Text("Page one"))
                    .PageBreak()
                    .Paragraph(p => p.Text("Page two"))
                    .ToBytes()));

        Assert.Contains("produced 2 page(s), but the approved baseline expects 1 page(s)", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RasterComparison_ReportsPixelDiffAndProducesDiffPng() {
        byte[] expected = VisualBaselineTestSupport.CreateRgbPng(2, 1, new byte[] {
            255, 255, 255,
            0, 0, 0
        });
        byte[] actual = VisualBaselineTestSupport.CreateRgbPng(2, 1, new byte[] {
            255, 255, 255,
            255, 0, 0
        });

        VisualRasterComparison comparison = CompareRasterImages(expected, actual, channelTolerance: 0, allowedDifferentPixels: 0);

        Assert.False(comparison.Passed);
        Assert.Equal(1, comparison.DifferentPixels);
        Assert.Equal(2, comparison.TotalPixels);
        Assert.Equal(255, comparison.MaxChannelDelta);
        Assert.True(comparison.DiffPng.Length > 0);
        Assert.Equal(2, VisualBaselineTestSupport.DecodePng(comparison.DiffPng, "PDF diff PNG is not a supported PNG file.").Width);
    }


}
