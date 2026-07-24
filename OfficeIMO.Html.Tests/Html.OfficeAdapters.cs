using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlOfficeAdapters {
    private static readonly byte[] OnePixelPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAABFSURBVEhLY1BNfv2flpgBXYDaeBhaILCzkSKMbt6oBRgY3bxRCzAwunmjFmBgdPNGLcDA6OaNWoCB0c3DsIDaeNQCghgAFxBXzP1LTe4AAAAASUVORK5CYII=");

    [Fact]
    public void ExcelHtml_ExportsSemanticWorksheetRichContent() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("Sales");
        sheet.CellValue(1, 1, "Region");
        sheet.CellValue(1, 2, "Amount");
        sheet.CellValue(2, 1, "North");
        sheet.CellValue(2, 2, 123.45);
        sheet.CellValue(3, 1, "South");
        sheet.CellValue(3, 2, 98.20);
        sheet.CellValue(4, 1, "West");
        sheet.CellValue(4, 2, 140.00);
        sheet.CellFormula(5, 2, "SUM(B2:B4)");
        sheet.SetComment(2, 2, "Reviewed with finance", "OfficeIMO");
        sheet.AddChartFromRange("A1:B4", row: 1, column: 5, widthPixels: 280, heightPixels: 180, type: ExcelChartType.ColumnClustered, title: "Revenue Trend");
        sheet.AddImage(7, 1, OnePixelPng, widthPixels: 48, heightPixels: 48, offsetXPixels: 9, offsetYPixels: 11, name: "Status Logo", altText: "Inline status marker");

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables,
            Theme = OfficeVisualThemeKind.Report
        });

        Assert.Contains("data-officeimo-profile=\"ExcelSemanticTables\"", html, StringComparison.Ordinal);
        Assert.Contains(">Region</th>", html, StringComparison.Ordinal);
        Assert.Contains(">North</td>", html, StringComparison.Ordinal);
        Assert.Contains("123.45", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-value-kind=\"number\" data-officeimo-value=\"123.45\"", html, StringComparison.Ordinal);
        Assert.Contains("--officeimo-accent", html, StringComparison.Ordinal);
        Assert.Contains("officeimo-formulas", html, StringComparison.Ordinal);
        Assert.Contains("SUM(B2:B4)", html, StringComparison.Ordinal);
        Assert.Contains("Reviewed with finance", html, StringComparison.Ordinal);
        Assert.Contains("Revenue Trend", html, StringComparison.Ordinal);
        Assert.Contains("officeimo-chart-data", html, StringComparison.Ordinal);
        Assert.Contains(">South</th>", html, StringComparison.Ordinal);
        Assert.Contains(">Amount</th>", html, StringComparison.Ordinal);
        Assert.Contains("Inline status marker", html, StringComparison.Ordinal);
        Assert.Contains("Offset: 9, 11", html, StringComparison.Ordinal);
        Assert.Contains("data:image/png;base64", html, StringComparison.Ordinal);
    }

    [Fact]
    public void ExcelHtml_ExportsVisualReviewFromSharedSvgRenderer() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("Visual");
        sheet.CellValue(1, 1, "Metric");
        sheet.CellValue(1, 2, "Score");
        sheet.CellValue(2, 1, "Ready");
        sheet.CellValue(2, 2, 12);
        sheet.CellValue(3, 1, "Done");
        sheet.CellValue(3, 2, 18);
        sheet.SetComment(2, 2, "Visual comment proof", "OfficeIMO");
        sheet.AddChartFromRange("A1:B3", row: 1, column: 4, widthPixels: 240, heightPixels: 150, type: ExcelChartType.ColumnClustered, title: "Visual Chart");
        sheet.AddImage(3, 3, OnePixelPng, widthPixels: 36, heightPixels: 36, name: "Visual Logo", altText: "Visual image marker");

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelVisualReview
        });

        Assert.Contains("data-officeimo-profile=\"ExcelVisualReview\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-visual-owner=\"OfficeIMO.Drawing\"", html, StringComparison.Ordinal);
        Assert.Contains("<svg", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Ready", html, StringComparison.Ordinal);
        Assert.Contains("Visual Chart", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-visual-proof=\"comment-callout\"", html, StringComparison.Ordinal);
        Assert.Contains("Visual comment proof", html, StringComparison.Ordinal);
        Assert.Contains("dependency-free callout approximation", html, StringComparison.Ordinal);
    }

    [Fact]
    public void ExcelHtml_VisualWorkbookNamespacesEmbeddedSvgIdsPerWorksheet() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet first = workbook.AddWorksheet("One");
        first.CellValue(1, 1, "Shared label");
        first.SetColumnWidth(1, 14);

        ExcelSheet second = workbook.AddWorksheet("Two");
        second.CellValue(1, 1, "Shared label");
        second.SetColumnWidth(1, 14);

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelVisualReview
        });

        Assert.Contains("id=\"officeimo-sheet-svg-one-1-xl-text-1-1\"", html, StringComparison.Ordinal);
        Assert.Contains("id=\"officeimo-sheet-svg-two-2-xl-text-1-1\"", html, StringComparison.Ordinal);
        Assert.Contains("clip-path=\"url(#officeimo-sheet-svg-one-1-xl-text-1-1)\"", html, StringComparison.Ordinal);
        Assert.Contains("clip-path=\"url(#officeimo-sheet-svg-two-2-xl-text-1-1)\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("id=\"xl-text-1-1\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("url(#xl-text-1-1)", html, StringComparison.Ordinal);
    }

    [Fact]
    public void ExcelHtml_CapabilityGalleryWritesSharedManifestForRichWorkbook() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("Gallery");
        sheet.CellValue(1, 1, "Region");
        sheet.CellValue(1, 2, "Amount");
        sheet.CellValue(2, 1, "North");
        sheet.CellValue(2, 2, 42);
        sheet.CellValue(3, 1, "South");
        sheet.CellValue(3, 2, 57);
        sheet.CellFormula(4, 2, "SUM(B2:B3)");
        sheet.SetComment(2, 2, "Gallery comment", "OfficeIMO");
        sheet.AddChartFromRange("A1:B3", row: 1, column: 4, widthPixels: 240, heightPixels: 150, type: ExcelChartType.ColumnClustered, title: "Gallery Chart");
        sheet.AddImage(6, 1, OnePixelPng, widthPixels: 36, heightPixels: 36, name: "Gallery Logo", altText: "Gallery image");

        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.HtmlOfficeAdapters", Guid.NewGuid().ToString("N"));
        HtmlCapabilityGalleryManifest manifest = workbook.SaveHtmlCapabilityGallery(directory, new ExcelHtmlCapabilityGalleryOptions {
            ScenarioId = "excel-gallery-rich",
            Title = "Excel Gallery Rich"
        });

        Assert.Equal("excel-gallery-rich", manifest.Result.Scenario.Id);
        Assert.Equal(HtmlConversionProfile.PositionedReview, manifest.Profile);
        Assert.Equal(new[] {
            OfficeHtmlConversionProfile.ExcelSemanticTables,
            OfficeHtmlConversionProfile.ExcelVisualReview
        }, manifest.OfficeProfiles);
        Assert.Equal(2, manifest.Result.Artifacts.Count);
        Assert.Contains(manifest.Expectations, expectation => expectation.Feature == "formulas" && expectation.Outcome == HtmlCapabilityGalleryExpectationOutcome.Preserved);
        Assert.Contains(manifest.Expectations, expectation => expectation.Feature == "comments" && expectation.Outcome == HtmlCapabilityGalleryExpectationOutcome.VisualProof);
        Assert.Contains(manifest.Expectations, expectation => expectation.Feature == "charts" && expectation.Outcome == HtmlCapabilityGalleryExpectationOutcome.Preserved);
        Assert.Contains(manifest.Expectations, expectation => expectation.Feature == "images" && expectation.Outcome == HtmlCapabilityGalleryExpectationOutcome.VisualProof);
        Assert.Contains(manifest.Result.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == "ExcelCommentVisualReviewRendered");
        Assert.Contains(manifest.Result.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == "ExcelChartSemanticDataPreserved");

        string semanticPath = Path.Combine(directory, "excel-gallery-rich.semantic.html");
        string visualPath = Path.Combine(directory, "excel-gallery-rich.visual.html");
        string manifestJsonPath = Path.Combine(directory, "excel-gallery-rich.manifest.json");
        Assert.True(File.Exists(semanticPath));
        Assert.True(File.Exists(visualPath));
        Assert.True(File.Exists(manifestJsonPath));
        string semanticHtml = File.ReadAllText(semanticPath);
        Assert.Contains("Gallery comment", semanticHtml, StringComparison.Ordinal);
        Assert.Contains("officeimo-chart-data", semanticHtml, StringComparison.Ordinal);
        Assert.Contains(">South</th>", semanticHtml, StringComparison.Ordinal);
        Assert.Contains(">57</td>", semanticHtml, StringComparison.Ordinal);
        string visualHtml = File.ReadAllText(visualPath);
        Assert.Contains("Gallery Chart", visualHtml, StringComparison.Ordinal);
        Assert.Contains("Gallery comment", visualHtml, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-visual-proof=\"comment-callout\"", visualHtml, StringComparison.Ordinal);

        string manifestJson = File.ReadAllText(manifestJsonPath);
        using JsonDocument json = JsonDocument.Parse(manifestJson);
        JsonElement root = json.RootElement;
        Assert.Equal("officeimo.html.capability-gallery", root.GetProperty("schemaId").GetString());
        Assert.Equal("excel-gallery-rich", root.GetProperty("scenario").GetProperty("id").GetString());
        Assert.Equal("PositionedReview", root.GetProperty("profile").GetProperty("id").GetString());
        JsonElement officeProfiles = root.GetProperty("officeProfiles");
        Assert.Equal(2, officeProfiles.GetArrayLength());
        Assert.Equal("ExcelSemanticTables", officeProfiles[0].GetProperty("id").GetString());
        Assert.Equal("ExcelVisualReview", officeProfiles[1].GetProperty("id").GetString());
        Assert.Equal("OfficeIMO.Drawing", officeProfiles[1].GetProperty("visualPrimitiveOwner").GetString());
        Assert.Equal(2, root.GetProperty("artifacts").GetArrayLength());
        Assert.Contains("charts", manifestJson, StringComparison.Ordinal);
        Assert.Contains("visual renderer owner", manifestJson, StringComparison.Ordinal);
        Assert.Contains("ExcelCommentVisualReviewRendered", manifestJson, StringComparison.Ordinal);
        Assert.Contains("ExcelChartSemanticDataPreserved", manifestJson, StringComparison.Ordinal);
    }

    [Fact]
    public void ExcelHtml_LoadsSemanticRichWorkbookBackToNativeWorkbook() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("Roundtrip");
        sheet.CellValue(1, 1, "Region");
        sheet.CellValue(1, 2, "Amount");
        sheet.CellValue(2, 1, "North");
        sheet.CellValue(2, 2, 123.45);
        sheet.CellValue(3, 1, "South");
        sheet.CellValue(3, 2, 98.20);
        sheet.CellValue(4, 1, "West");
        sheet.CellValue(4, 2, 140.00);
        sheet.CellFormula(5, 2, "SUM(B2:B4)");
        sheet.SetComment(2, 2, "Reviewed with finance", "OfficeIMO");
        sheet.AddChartFromRange("A1:B4", row: 1, column: 5, widthPixels: 280, heightPixels: 180, type: ExcelChartType.ColumnClustered, title: "Revenue Trend");
        sheet.AddImage(7, 1, OnePixelPng, widthPixels: 48, heightPixels: 48, offsetXPixels: 9, offsetYPixels: 11, name: "Status Logo", altText: "Inline status marker");

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables,
            Theme = OfficeVisualThemeKind.Report
        });

        Assert.Contains("Cell: 1, 5", html, StringComparison.Ordinal);
        Assert.Contains("Size: 280x180", html, StringComparison.Ordinal);
        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(
            html,
            HtmlConversionDocumentOptions.CreateTrustedProfile()).ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;
        ExcelSheet importedSheet = imported.Sheets.Single(importedSheet => importedSheet.Name == "Roundtrip");

        Assert.Equal(1, result.Sheets);
        Assert.True(result.Cells >= 8);
        Assert.Equal(1, result.Formulas);
        Assert.Equal(1, result.Comments);
        Assert.Equal(1, result.Images);
        Assert.Equal(1, result.Charts);
        Assert.Empty(result.Report.Diagnostics);
        Assert.True(importedSheet.TryGetCellText(2, 1, out string region));
        Assert.Equal("North", region);
        Assert.Contains(importedSheet.GetFormulaCells(), formula => formula.CellReference == "B5" && formula.Formula == "SUM(B2:B4)");
        Assert.Contains(importedSheet.GetComments(), comment => comment.CellReference == "B2" && comment.Text == "Reviewed with finance" && comment.Author == "OfficeIMO");
        ExcelImage importedImage = Assert.Single(importedSheet.Images);
        Assert.Equal(9, importedImage.OffsetXPixels);
        Assert.Equal(11, importedImage.OffsetYPixels);
        ExcelChart importedChart = Assert.Single(importedSheet.Charts);
        Assert.True(importedChart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        Assert.Equal(1, snapshot.RowIndex);
        Assert.Equal(5, snapshot.ColumnIndex);
        Assert.Equal(280, snapshot.WidthPixels);
        Assert.Equal(180, snapshot.HeightPixels);
        Assert.Equal(new[] { "North", "South", "West" }, snapshot.Data.Categories);
        ExcelChartSeries importedSeries = Assert.Single(snapshot.Data.Series);
        Assert.Equal("Amount", importedSeries.Name);
        Assert.Equal(new[] { 123.45D, 98.2D, 140D }, importedSeries.Values);

        string roundTripHtml = imported.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables,
            Theme = OfficeVisualThemeKind.Report
        });

        Assert.Contains("data-officeimo-profile=\"ExcelSemanticTables\"", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("North", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("SUM(B2:B4)", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("Reviewed with finance", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("Revenue Trend", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("officeimo-chart-data", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains(">South</th>", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains(">Amount</th>", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("Inline status marker", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("data:image/png;base64", roundTripHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void ExcelHtml_RoundTripsImageTransformsAbsoluteAnchorAndDrawingOrder() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("Drawings");
        sheet.CellValue(1, 1, "Category");
        sheet.CellValue(1, 2, "Value");
        sheet.CellValue(2, 1, "A");
        sheet.CellValue(2, 2, 10);
        sheet.CellValue(3, 1, "B");
        sheet.CellValue(3, 2, 20);
        sheet.AddChartFromRange("A1:B3", row: 1, column: 4, widthPixels: 240, heightPixels: 140, type: ExcelChartType.ColumnClustered, title: "Layer base");
        ExcelImage image = sheet.AddImageAbsolute(33, 44, OnePixelPng, widthPixels: 52, heightPixels: 48, name: "Layer image", altText: "Absolute layer");
        image.SetRotation(17.5).SetFlip(horizontal: true, vertical: true).SetCropRatio(0.125D, 0.25D, 0.0625D, 0.1875D);

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables
        });

        Assert.Contains("data-officeimo-anchor=\"absolute\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-x=\"33\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-y=\"44\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-layer-kind=\"image\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-layer-kind=\"chart\"", html, StringComparison.Ordinal);
        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;
        ExcelSheet importedSheet = imported.Sheets.Single(importedSheet => importedSheet.Name == "Drawings");

        Assert.Equal(1, result.Images);
        Assert.Equal(1, result.Charts);
        Assert.Empty(result.Report.Diagnostics);
        ExcelChart importedChart = Assert.Single(importedSheet.Charts);
        ExcelImage importedImage = Assert.Single(importedSheet.Images);
        Assert.True(importedImage.HasAbsoluteAnchor);
        Assert.True(importedImage.TryGetAbsoluteAnchorBounds(out int x, out int y, out int width, out int height));
        Assert.Equal(33, x);
        Assert.Equal(44, y);
        Assert.Equal(52, width);
        Assert.Equal(48, height);
        Assert.Equal(17.5D, importedImage.RotationDegrees, 3);
        Assert.True(importedImage.FlipHorizontal);
        Assert.True(importedImage.FlipVertical);
        Assert.Equal(0.125D, importedImage.CropLeftRatio, 3);
        Assert.Equal(0.25D, importedImage.CropTopRatio, 3);
        Assert.Equal(0.0625D, importedImage.CropRightRatio, 3);
        Assert.Equal(0.1875D, importedImage.CropBottomRatio, 3);
        Assert.True(importedChart.DrawingOrder < importedImage.DrawingOrder);
    }

    [Fact]
    public void ExcelHtml_RejectsCropPairsThatConsumeTheWholeImage() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("Drawings");
        sheet.CellValue(1, 1, "Seed");
        sheet.AddImageAbsolute(10, 10, OnePixelPng, widthPixels: 40, heightPixels: 30)
            .SetCropRatio(0.1D, 0D, 0.1D, 0D);
        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables
        });
        html = System.Text.RegularExpressions.Regex.Replace(
            html, "data-officeimo-crop-left=\"[^\"]*\"", "data-officeimo-crop-left=\"0.75\"");
        html = System.Text.RegularExpressions.Regex.Replace(
            html, "data-officeimo-crop-right=\"[^\"]*\"", "data-officeimo-crop-right=\"0.75\"");

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;
        ExcelImage image = Assert.Single(imported.Sheets.Single().Images);

        Assert.Equal(0D, image.CropLeftRatio);
        Assert.Equal(0D, image.CropRightRatio);
        Assert.Contains(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticValueInvalid);
    }

    [Fact]
    public void ExcelHtml_RoundTripsTwoCellImageAnchors() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("Anchors");
        sheet.CellValue(1, 1, "Seed");
        ExcelImage image = sheet.AddImageToRange("B2:D5", OnePixelPng, "image/png", offsetXPixels: 7, offsetYPixels: 8, endOffsetXPixels: 9, endOffsetYPixels: 10, name: "Range image", altText: "Two-cell anchor");
        image.SetRotation(12.5D).SetFlip(horizontal: true, vertical: false);

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables
        });

        Assert.Contains("data-officeimo-anchor=\"twoCell\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-to-row=\"6\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-to-column=\"5\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-to-offset-x=\"9\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-to-offset-y=\"10\"", html, StringComparison.Ordinal);
        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;
        ExcelSheet importedSheet = imported.Sheets.Single(importedSheet => importedSheet.Name == "Anchors");

        Assert.Equal(1, result.Images);
        Assert.Empty(result.Report.Diagnostics);
        ExcelImage importedImage = Assert.Single(importedSheet.Images);
        Assert.True(importedImage.HasTwoCellAnchor);
        Assert.Equal(2, importedImage.RowIndex);
        Assert.Equal(2, importedImage.ColumnIndex);
        Assert.Equal(7, importedImage.OffsetXPixels);
        Assert.Equal(8, importedImage.OffsetYPixels);
        Assert.Equal(6, importedImage.ToRowIndex);
        Assert.Equal(5, importedImage.ToColumnIndex);
        Assert.Equal(9, importedImage.ToOffsetXPixels);
        Assert.Equal(10, importedImage.ToOffsetYPixels);
        Assert.Equal(12.5D, importedImage.RotationDegrees, 3);
        Assert.True(importedImage.FlipHorizontal);
    }

    [Fact]
    public void ExcelHtml_RoundTripsSameCellTwoCellImageMarkers() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("SameCellAnchor");
        sheet.CellValue(1, 1, "Seed");
        sheet.AddImageToRange("B2:B2", OnePixelPng, "image/png", name: "Same-cell image", altText: "Same-cell anchor")
            .SetTwoCellEndingMarker(2, 2, 3, 4);

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables
        });

        Assert.Contains("data-officeimo-to-row=\"2\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-to-column=\"2\"", html, StringComparison.Ordinal);
        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;
        ExcelSheet importedSheet = imported.Sheets.Single(importedSheet => importedSheet.Name == "SameCellAnchor");

        ExcelImage importedImage = Assert.Single(importedSheet.Images);
        Assert.True(importedImage.HasTwoCellAnchor);
        Assert.Equal(2, importedImage.RowIndex);
        Assert.Equal(2, importedImage.ColumnIndex);
        Assert.Equal(2, importedImage.ToRowIndex);
        Assert.Equal(2, importedImage.ToColumnIndex);
        Assert.Equal(3, importedImage.ToOffsetXPixels);
        Assert.Equal(4, importedImage.ToOffsetYPixels);
    }

    [Fact]
    public void ExcelHtml_LoadHonorsSemanticRangeOrigin() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("Offset");
        sheet.CellValue(2, 2, "Region");
        sheet.CellValue(2, 3, "Amount");
        sheet.CellValue(3, 2, "North");
        sheet.CellValue(3, 3, 123);
        sheet.SetComment(3, 3, "Aligned comment", "OfficeIMO");

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables
        });

        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;
        ExcelSheet importedSheet = imported.Sheets.Single(importedSheet => importedSheet.Name == "Offset");

        Assert.True(importedSheet.TryGetCellText(2, 2, out string header));
        Assert.Equal("Region", header);
        Assert.True(importedSheet.TryGetCellText(3, 3, out string amount));
        Assert.Equal("123", amount);
        Assert.False(importedSheet.TryGetCellText(1, 1, out _));
        Assert.Contains(importedSheet.GetComments(), comment => comment.CellReference == "C3" && comment.Text == "Aligned comment");
    }

    [Fact]
    public void ExcelHtml_ExportsEmptySheetWithoutCreatingPlaceholderCell() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        workbook.AddWorksheet("Empty");

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables,
            EmptyCellText = "(empty)"
        });

        Assert.Contains("No used cells.", html, StringComparison.Ordinal);
        Assert.DoesNotContain("(empty)", html, StringComparison.Ordinal);

        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;
        ExcelSheet importedSheet = Assert.Single(imported.Sheets);

        Assert.Equal("Empty", importedSheet.Name);
        Assert.Equal(0, result.Cells);
        Assert.False(importedSheet.TryGetCellText(1, 1, out _));
    }

    [Fact]
    public void ExcelHtml_ReportsTruncationWhenRetainedWindowContainsNoUsedCells() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("Sparse");
        sheet.CellValue(1, A1.MaxColumns, "Far column");
        sheet.CellValue(A1.MaxRows, 1, "Far row");

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables,
            MaxRowsPerSheet = 2,
            MaxColumnsPerSheet = 2
        });

        Assert.Contains("No cells within export limits.", html, StringComparison.Ordinal);
        Assert.Contains("Columns truncated: 2 of 16384 exported.", html, StringComparison.Ordinal);
        Assert.Contains("Rows truncated: 2 of 1048576 exported.", html, StringComparison.Ordinal);
        Assert.DoesNotContain("No used cells.", html, StringComparison.Ordinal);
    }

    [Fact]
    public void ExcelHtml_RoundTripsHiddenWorksheetVisibility() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet visible = workbook.AddWorksheet("Visible");
        ExcelSheet hidden = workbook.AddWorksheet("Hidden");
        ExcelSheet veryHidden = workbook.AddWorksheet("VeryHidden");
        visible.CellValue(1, 1, "Visible");
        hidden.CellValue(1, 1, "Hidden");
        veryHidden.CellValue(1, 1, "Very hidden");
        hidden.SetHidden(true);
        veryHidden.SetVeryHidden(true);

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables
        });

        Assert.Contains("data-officeimo-sheet=\"Hidden\" data-officeimo-range=\"A1:A1\" data-officeimo-visibility=\"hidden\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-sheet=\"VeryHidden\" data-officeimo-range=\"A1:A1\" data-officeimo-visibility=\"veryHidden\"", html, StringComparison.Ordinal);

        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;
        ExcelSheet importedVisible = imported.Sheets.Single(sheet => sheet.Name == "Visible");
        ExcelSheet importedHidden = imported.Sheets.Single(sheet => sheet.Name == "Hidden");
        ExcelSheet importedVeryHidden = imported.Sheets.Single(sheet => sheet.Name == "VeryHidden");

        Assert.False(importedVisible.Hidden);
        Assert.False(importedVisible.VeryHidden);
        Assert.True(importedHidden.Hidden);
        Assert.False(importedHidden.VeryHidden);
        Assert.True(importedVeryHidden.VeryHidden);
    }

    [Fact]
    public void ExcelHtml_DoesNotImportEmptyCellTextPlaceholdersAsValues() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("Placeholders");
        sheet.CellValue(1, 1, "Left");
        sheet.CellValue(1, 3, "Right");

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables,
            EmptyCellText = "(empty)"
        });

        Assert.Contains("data-officeimo-empty=\"true\">(empty)", html, StringComparison.Ordinal);

        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;
        ExcelSheet importedSheet = Assert.Single(imported.Sheets);

        Assert.Equal(2, result.Cells);
        Assert.False(importedSheet.TryGetCellValueSnapshot(1, 2, out _));
    }

    [Fact]
    public void ExcelHtml_LoadCreatesValidWorkbookWhenNoSheetSectionsExist() {
        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse("<main><p>No workbook markup</p></main>").ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;

        ExcelSheet importedSheet = Assert.Single(imported.Sheets);
        using var output = new MemoryStream();
        imported.Save(output);

        Assert.Equal("Imported", importedSheet.Name);
        Assert.NotEmpty(output.ToArray());
        HtmlDiagnostic diagnostic = Assert.Single(result.Report.Diagnostics);
        Assert.Equal(HtmlConversionDiagnosticCodes.SemanticContentMissing, diagnostic.Code);
        Assert.Equal(HtmlDiagnosticSeverity.Error, diagnostic.Severity);
        Assert.Equal(HtmlConversionLossKind.Failure, diagnostic.LossKind);
        Assert.Contains("No semantic Excel sheet sections", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ExcelHtml_LoadPreservesFormulaWhitespaceFromSemanticInventory() {
        string formula = "CONCAT(\"A  B\",\" C\")";
        string html = """
            <main>
              <section class="officeimo-sheet" data-officeimo-sheet="Formulas" data-officeimo-range="A1:A1">
                <table><tr><td>seed</td></tr></table>
                <section class="officeimo-feature officeimo-formulas">
                  <ul class="officeimo-feature-list">
                    <li class="officeimo-feature-item" data-officeimo-cell="A2">
                      <span class="officeimo-feature-label">A2</span><code>CONCAT("A  B"," C")</code>
                    </li>
                  </ul>
                </section>
              </section>
            </main>
            """;

        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(
            html,
            HtmlConversionDocumentOptions.CreateTrustedProfile()).ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;
        ExcelSheet importedSheet = Assert.Single(imported.Sheets);

        Assert.Contains(importedSheet.GetFormulaCells(), cell => cell.CellReference == "A2" && cell.Formula == formula);
    }

    [Fact]
    public void ExcelHtml_LoadClearsFallbackTextWhenApplyingSemanticFormula() {
        string html = """
            <main>
              <section class="officeimo-sheet" data-officeimo-sheet="Formulas" data-officeimo-range="A1:A3">
                <table>
                  <tr><td data-officeimo-value-kind="number" data-officeimo-value="1">1</td></tr>
                  <tr><td data-officeimo-value-kind="number" data-officeimo-value="2">2</td></tr>
                  <tr><td data-officeimo-value-kind="formula" data-officeimo-value="SUM(A1:A2)">3</td></tr>
                </table>
                <section class="officeimo-feature officeimo-formulas">
                  <ul class="officeimo-feature-list">
                    <li class="officeimo-feature-item" data-officeimo-cell="A3"><code>SUM(A1:A2)</code></li>
                  </ul>
                </section>
              </section>
            </main>
            """;

        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(
            html,
            HtmlConversionDocumentOptions.CreateTrustedProfile()).ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;
        ExcelSheet importedSheet = Assert.Single(imported.Sheets);

        Assert.True(importedSheet.TryGetCellValueSnapshot(3, 1, out ExcelCellValueSnapshot? snapshot));
        Assert.Equal(ExcelCellValueKind.Formula, snapshot!.Kind);
        Assert.Equal("SUM(A1:A2)", snapshot.RawValue);
    }

    [Fact]
    public void ExcelHtml_RoundTripsSemanticCellValueKinds() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("Typed");
        sheet.CellValue(1, 1, "Label");
        sheet.CellValue(1, 2, "Value");
        sheet.CellValue(2, 1, "Numeric text");
        sheet.CellValue(2, 2, "123");
        sheet.CellValue(3, 1, "Amount");
        sheet.CellValue(3, 2, 123.45D);
        sheet.CellValue(4, 1, "Flag");
        sheet.CellValue(4, 2, true);

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables
        });

        Assert.Contains("data-officeimo-value-kind=\"number\" data-officeimo-value=\"123.45\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-value-kind=\"boolean\" data-officeimo-value=\"1\"", html, StringComparison.Ordinal);

        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;
        ExcelSheet importedSheet = Assert.Single(imported.Sheets);

        Assert.True(importedSheet.TryGetCellValueSnapshot(2, 2, out ExcelCellValueSnapshot? textSnapshot));
        Assert.Equal(ExcelCellValueKind.Text, textSnapshot!.Kind);
        Assert.True(importedSheet.TryGetCellValueSnapshot(3, 2, out ExcelCellValueSnapshot? numberSnapshot));
        Assert.Equal(ExcelCellValueKind.Number, numberSnapshot!.Kind);
        Assert.Equal("123.45", numberSnapshot.RawValue);
        Assert.True(importedSheet.TryGetCellValueSnapshot(4, 2, out ExcelCellValueSnapshot? boolSnapshot));
        Assert.Equal(ExcelCellValueKind.Boolean, boolSnapshot!.Kind);
        Assert.Equal("1", boolSnapshot.RawValue);
    }

    [Fact]
    public void ExcelHtml_RoundTripsSemanticErrorCellValueKind() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("Errors");
        sheet.CellValue(1, 1, "Error");
        sheet.CellError(2, 1, "#DIV/0!");

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables
        });

        Assert.Contains("data-officeimo-value-kind=\"error\" data-officeimo-value=\"#DIV/0!\"", html, StringComparison.Ordinal);

        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;
        ExcelSheet importedSheet = Assert.Single(imported.Sheets);

        Assert.True(importedSheet.TryGetCellValueSnapshot(2, 1, out ExcelCellValueSnapshot? snapshot));
        Assert.Equal(ExcelCellValueKind.Error, snapshot!.Kind);
        Assert.Equal("#DIV/0!", snapshot.RawValue);
    }

    [Fact]
    public void ExcelHtml_MaxRowsPerSheetFiltersFeatureInventoryToExportedRows() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("Limited");
        sheet.CellValue(1, 1, "Name");
        sheet.CellValue(1, 2, "Amount");
        sheet.CellValue(2, 1, "Visible");
        sheet.CellFormula(2, 2, "1+1");
        sheet.SetComment(2, 1, "Visible comment", "OfficeIMO");
        sheet.AddChartFromRange("A1:B2", row: 2, column: 4, widthPixels: 160, heightPixels: 90, type: ExcelChartType.ColumnClustered, title: "Visible Chart");
        sheet.AddImage(2, 3, OnePixelPng, widthPixels: 24, heightPixels: 24, name: "Visible Image", altText: "Visible image");
        sheet.CellValue(4, 1, "Truncated");
        sheet.CellFormula(4, 2, "99+1");
        sheet.CellValue(5, 1, "Also truncated");
        sheet.CellValue(5, 2, 25);
        sheet.SetComment(4, 1, "Truncated comment", "OfficeIMO");
        sheet.AddChartFromRange("A4:B5", row: 4, column: 4, widthPixels: 160, heightPixels: 90, type: ExcelChartType.ColumnClustered, title: "Truncated Chart");
        sheet.AddImage(4, 3, OnePixelPng, widthPixels: 24, heightPixels: 24, name: "Truncated Image", altText: "Truncated image");
        sheet.AddImageAbsolute(0, 2000, OnePixelPng, widthPixels: 24, heightPixels: 24, name: "Truncated Absolute Image", altText: "Truncated absolute image");

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables,
            MaxRowsPerSheet = 2
        });

        Assert.Contains("Rows truncated: 2 of 5 exported.", html, StringComparison.Ordinal);
        Assert.Contains("1+1", html, StringComparison.Ordinal);
        Assert.Contains("Visible comment", html, StringComparison.Ordinal);
        Assert.Contains("Visible Chart", html, StringComparison.Ordinal);
        Assert.Contains("Visible image", html, StringComparison.Ordinal);
        Assert.DoesNotContain("99+1", html, StringComparison.Ordinal);
        Assert.DoesNotContain("Truncated comment", html, StringComparison.Ordinal);
        Assert.DoesNotContain("Truncated Chart", html, StringComparison.Ordinal);
        Assert.DoesNotContain("Truncated image", html, StringComparison.Ordinal);
        Assert.DoesNotContain("Truncated absolute image", html, StringComparison.Ordinal);
    }

    [Fact]
    public void ExcelHtml_LoadPreservesRawTextCellWhitespace() {
        string value = "A  B\n C";
        string html = """
            <main>
              <section class="officeimo-sheet" data-officeimo-sheet="Whitespace" data-officeimo-range="A1:A1">
                <table>
                  <tr><td data-officeimo-value-kind="text" data-officeimo-value="A  B&#10; C">A B C</td></tr>
                </table>
              </section>
            </main>
            """;

        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;
        using var stream = new MemoryStream();
        imported.Save(stream);
        stream.Position = 0;
        using ExcelDocument persisted = ExcelDocument.Load(stream, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
        ExcelSheet importedSheet = Assert.Single(persisted.Sheets);

        Assert.True(importedSheet.TryGetCellValueSnapshot(1, 1, out ExcelCellValueSnapshot? snapshot));
        Assert.Equal(ExcelCellValueKind.Text, snapshot!.Kind);
        Assert.Equal(value, snapshot.Text);
    }

    [Fact]
    public void ExcelHtml_LoadPreservesCommentTextWhitespace() {
        string html = """
            <main>
              <section class="officeimo-sheet" data-officeimo-sheet="Comments" data-officeimo-range="A1:A1">
                <table><tr><td>seed</td></tr></table>
                <section class="officeimo-feature officeimo-comments">
                  <ul class="officeimo-feature-list">
                    <li class="officeimo-feature-item" data-officeimo-cell="A1">
                      <span class="officeimo-feature-label">A1</span>
                      <p>Line 1&#10;Line 2  with  spacing</p>
                    </li>
                  </ul>
                </section>
              </section>
            </main>
            """;

        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;
        var comment = Assert.Single(imported.Sheets[0].GetComments());

        Assert.Equal("Line 1\nLine 2  with  spacing", comment.Text);
    }

    [Fact]
    public void ExcelHtml_RoundTripsScatterChartXValuesInSemanticChartData() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("Scatter");
        var data = new ExcelChartData(
            new[] { "1.5", "2.5" },
            new[] { new ExcelChartSeries("Points", new[] { 10D, 20D }, new[] { 1.5D, 2.5D }, ExcelChartType.Scatter) });
        sheet.AddChart(data, row: 1, column: 1, widthPixels: 320, heightPixels: 180, type: ExcelChartType.Scatter, title: "Scatter");

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables
        });

        Assert.Contains("data-officeimo-x=\"1.5\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-x=\"2.5\"", html, StringComparison.Ordinal);

        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;
        ExcelChart chart = Assert.Single(imported.Sheets.Single(sheet => sheet.Name == "Scatter").Charts);
        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        ExcelChartSeries series = Assert.Single(snapshot.Data.Series);
        Assert.Equal(new[] { 1.5D, 2.5D }, series.XValues);
        Assert.Equal(new[] { 10D, 20D }, series.Values);
    }

    [Fact]
    public void ExcelHtml_RoundTripsVariableLengthScatterSeriesInSemanticChartData() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("Scatter");
        var data = new ExcelChartData(
            new[] { "1", "2", "3" },
            new[] {
                new ExcelChartSeries("Forecast", new[] { 10D, 20D, 30D }, new[] { 1D, 2D, 3D }, ExcelChartType.Scatter),
                new ExcelChartSeries("Outliers", new[] { 40D, 50D }, new[] { 4D, 5D }, ExcelChartType.Scatter)
            });
        sheet.AddChart(data, row: 1, column: 1, widthPixels: 320, heightPixels: 180, type: ExcelChartType.Scatter, title: "Scatter");

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables
        });

        Assert.Contains("data-officeimo-x=\"5\"", html, StringComparison.Ordinal);

        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;
        ExcelChart chart = Assert.Single(imported.Sheets.Single(sheet => sheet.Name == "Scatter").Charts);
        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        Assert.Equal(new[] { 1D, 2D, 3D }, snapshot.Data.Series[0].XValues);
        Assert.Equal(new[] { 10D, 20D, 30D }, snapshot.Data.Series[0].Values);
        Assert.Equal(new[] { 4D, 5D }, snapshot.Data.Series[1].XValues);
        Assert.Equal(new[] { 40D, 50D }, snapshot.Data.Series[1].Values);
    }

    [Fact]
    public void ExcelChart_ScatterSnapshotPrefersLiveRangesOverCachedValues() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("ScatterLive");
        var data = new ExcelChartData(
            new[] { "1", "2" },
            new[] { new ExcelChartSeries("Points", new[] { 10D, 20D }, new[] { 1D, 2D }, ExcelChartType.Scatter) });
        ExcelChart chart = sheet.AddChart(data, row: 1, column: 5, widthPixels: 320, heightPixels: 180, type: ExcelChartType.Scatter, title: "Scatter");
        ExcelChartDataRange range = chart.DataRange!;
        ExcelSheet dataSheet = workbook[range.SheetName];

        dataSheet.CellValue(range.CategoryStartRow, range.CategoryStartColumn, 3D);
        dataSheet.CellValue(range.CategoryStartRow + 1, range.CategoryStartColumn, 4D);
        dataSheet.CellValue(range.SeriesStartRow, range.SeriesStartColumn, 30D);
        dataSheet.CellValue(range.SeriesStartRow + 1, range.SeriesStartColumn, 40D);

        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        ExcelChartSeries series = Assert.Single(snapshot.Data.Series);
        Assert.Equal(new[] { 3D, 4D }, series.XValues);
        Assert.Equal(new[] { 30D, 40D }, series.Values);
    }

    [Fact]
    public void ExcelHtml_RoundTripsComboChartSeriesTypesInSemanticChartData() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("Combo");
        var data = new ExcelChartData(
            new[] { "Q1", "Q2" },
            new[] {
                new ExcelChartSeries("Sales", new[] { 10D, 20D }, ExcelChartType.ColumnClustered),
                new ExcelChartSeries("Trend", new[] { 12D, 22D }, ExcelChartType.Line)
            });
        sheet.AddChart(data, row: 1, column: 1, widthPixels: 320, heightPixels: 180, type: ExcelChartType.ColumnClustered, title: "Combo");

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables
        });

        Assert.Contains("data-officeimo-chart-type=\"ColumnClustered\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-chart-type=\"Line\"", html, StringComparison.Ordinal);

        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;
        ExcelChart chart = Assert.Single(imported.Sheets.Single(sheet => sheet.Name == "Combo").Charts);
        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        Assert.Collection(
            snapshot.Data.Series,
            series => Assert.Equal(ExcelChartType.ColumnClustered, series.ChartType),
            series => Assert.Equal(ExcelChartType.Line, series.ChartType));
    }

    [Fact]
    public void PowerPointHtml_RoundTripsScatterChartXValuesInSemanticChartData() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        var data = new OfficeChartData(new[] { "1.5", "2.5" }, new[] {
            new OfficeChartSeries("Forecast", new[] { 10D, 20D }, new[] { 1.5D, 2.5D })
        });
        slide.AddChartPoints(OfficeChartKind.Scatter, data, 72, 96, 240, 140).SetTitle("Scatter");

        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides
        });

        Assert.Contains("data-officeimo-x=\"1.5\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-x=\"2.5\"", html, StringComparison.Ordinal);

        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using PowerPointPresentation imported = result.Value;
        PowerPointChart importedChart = Assert.Single(imported.Slides[0].Charts);
        Assert.True(importedChart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        OfficeChartSeries series = Assert.Single(snapshot.Data.Series);
        Assert.Equal(new[] { 1.5D, 2.5D }, series.XValues);
        Assert.Equal(new[] { 10D, 20D }, series.Values);
    }

    [Fact]
    public void PowerPointHtml_RejectsXValuesOnMismatchedNonScatterSeries() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Actual", new[] { 10D, 20D })
        });
        slide.AddChartPoints(OfficeChartKind.ColumnClustered, data, 72, 96, 240, 140).SetTitle("Columns");

        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides
        });
        html = html.Replace("<td>10</td>", "<td data-officeimo-x=\"1\">10</td>", StringComparison.Ordinal)
            .Replace("<td>20</td>", string.Empty, StringComparison.Ordinal);

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using PowerPointPresentation imported = result.Value;
        PowerPointChart importedChart = Assert.Single(imported.Slides[0].Charts);

        Assert.True(importedChart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        OfficeChartSeries series = Assert.Single(snapshot.Data.Series);
        Assert.Equal(2, series.Values.Count);
        Assert.Null(series.XValues);
        Assert.Contains(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated);
    }

    [Fact]
    public void PowerPointHtml_RoundTripsVariableLengthScatterSeriesInSemanticChartData() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        var data = new OfficeChartData(new[] { "1", "2", "3" }, new[] {
            new OfficeChartSeries("Forecast", new[] { 10D, 20D, 30D }, new[] { 1D, 2D, 3D }),
            new OfficeChartSeries("Outliers", new[] { 40D, 50D }, new[] { 4D, 5D })
        });
        slide.AddChartPoints(OfficeChartKind.Scatter, data, 72, 96, 240, 140).SetTitle("Scatter");

        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides
        });

        Assert.Contains("data-officeimo-x=\"5\"", html, StringComparison.Ordinal);

        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using PowerPointPresentation imported = result.Value;
        PowerPointChart importedChart = Assert.Single(imported.Slides[0].Charts);
        Assert.True(importedChart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(new[] { 1D, 2D, 3D }, snapshot.Data.Series[0].XValues);
        Assert.Equal(new[] { 10D, 20D, 30D }, snapshot.Data.Series[0].Values);
        Assert.Equal(new[] { 4D, 5D }, snapshot.Data.Series[1].XValues);
        Assert.Equal(new[] { 40D, 50D }, snapshot.Data.Series[1].Values);
    }

    [Fact]
    public void PowerPointHtml_ExportsSemanticSlidesWithExtractionProof() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        slide.AddTitle("Roadmap");
        slide.AddTextBox("HTML end to end");
        slide.Notes.Text = "Presenter reminder";
        using (var image = new MemoryStream(OnePixelPng)) {
            PowerPointPicture picture = slide.AddPicturePoints(image, OfficeIMO.PowerPoint.ImagePartType.Png, 72, 140, 72, 72);
            picture.Name = "Architecture badge";
            picture.AltText = "Reusable renderer badge";
        }

        OfficeChartData chartData = new(
            new[] { "Q1", "Q2", "Q3" },
            new[] { new OfficeChartSeries("Actual", new[] { 10D, 18D, 24D }) });
        slide.AddChartPoints(OfficeChartKind.ColumnClustered, chartData, 180, 130, 260, 150).SetTitle("Pipeline");

        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides
        });

        Assert.Contains("data-officeimo-profile=\"PowerPointSemanticSlides\"", html, StringComparison.Ordinal);
        Assert.Contains(">Roadmap</p>", html, StringComparison.Ordinal);
        Assert.Contains("HTML end to end", html, StringComparison.Ordinal);
        Assert.Contains("Presenter reminder", html, StringComparison.Ordinal);
        Assert.Contains("officeimo-source-markdown", html, StringComparison.Ordinal);
        Assert.Contains("officeimo-images", html, StringComparison.Ordinal);
        Assert.Contains("Reusable renderer badge", html, StringComparison.Ordinal);
        Assert.Contains("data:image/png;base64", html, StringComparison.Ordinal);
        Assert.Contains("Pipeline", html, StringComparison.Ordinal);
        Assert.Contains("ClusteredColumn", html, StringComparison.Ordinal);
        Assert.Contains("officeimo-chart-data", html, StringComparison.Ordinal);
        Assert.Contains("<th>Q2</th>", html, StringComparison.Ordinal);
        Assert.Contains("<td>18</td>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerPointHtml_ExportsPositionedReviewSlides() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        slide.AddTextBoxPoints("Positioned", 72, 96, 240, 60);
        using (var image = new MemoryStream(OnePixelPng)) {
            PowerPointPicture picture = slide.AddPicturePoints(image, OfficeIMO.PowerPoint.ImagePartType.Png, 72, 180, 72, 72);
            picture.AltText = "Positioned image";
        }

        OfficeChartData chartData = new(
            new[] { "Q1", "Q2", "Q3" },
            new[] { new OfficeChartSeries("Actual", new[] { 8D, 13D, 21D }) });
        slide.AddChartPoints(OfficeChartKind.ColumnClustered, chartData, 180, 180, 260, 150).SetTitle("Positioned Pipeline");

        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointVisualReview
        });

        Assert.Contains("data-officeimo-profile=\"PowerPointVisualReview\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-visual-boundary=\"positioned-review\"", html, StringComparison.Ordinal);
        Assert.Contains("left:72pt;top:96pt;width:240pt;height:60pt", html, StringComparison.Ordinal);
        Assert.Contains("Positioned", html, StringComparison.Ordinal);
        Assert.Contains("data:image/png;base64", html, StringComparison.Ordinal);
        Assert.Contains("officeimo-chart-rendered", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-visual-owner=\"OfficeIMO.Drawing\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-chart-kind=\"ColumnClustered\"", html, StringComparison.Ordinal);
        Assert.Contains("<svg", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Positioned Pipeline", html, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerPointHtml_SemanticSlidesSkipHiddenShapesByDefault() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        slide.AddTextBox("Visible briefing");
        PowerPointTextBox hidden = slide.AddTextBox("Hidden briefing");
        hidden.Hidden = true;
        using (var image = new MemoryStream(OnePixelPng)) {
            PowerPointPicture hiddenPicture = slide.AddPicturePoints(image, OfficeIMO.PowerPoint.ImagePartType.Png, 80, 90, 120, 72);
            hiddenPicture.Name = "Hidden image marker";
            hiddenPicture.Hidden = true;
        }

        OfficeChartData hiddenChartData = new(
            new[] { "Q1", "Q2" },
            new[] { new OfficeChartSeries("Hidden", new[] { 10D, 18D }) });
        PowerPointChart hiddenChart = slide.AddChartPoints(OfficeChartKind.ColumnClustered, hiddenChartData, 120, 90, 240, 140);
        hiddenChart.SetTitle("Hidden chart marker");
        hiddenChart.Hidden = true;

        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides
        });

        Assert.Contains("Visible briefing", html, StringComparison.Ordinal);
        Assert.DoesNotContain("Hidden briefing", html, StringComparison.Ordinal);
        Assert.DoesNotContain("Hidden image marker", html, StringComparison.Ordinal);
        Assert.DoesNotContain("Hidden chart marker", html, StringComparison.Ordinal);

        string htmlWithHidden = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides,
            IncludeHiddenShapes = true
        });

        Assert.Contains("Hidden briefing", htmlWithHidden, StringComparison.Ordinal);
        Assert.Contains("Hidden image marker", htmlWithHidden, StringComparison.Ordinal);
        Assert.Contains("Hidden chart marker", htmlWithHidden, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerPointHtml_VisualReviewIncludesInheritedLayoutShapes() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointTextBox layoutText = slide.AddTextBoxPoints("Layout footer", 24, 200, 200, 24);
        P.Shape slideShape = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
            .Elements<P.Shape>()
            .Single(shape => shape.InnerText.Contains("Layout footer", StringComparison.Ordinal));
        P.Shape layoutShape = (P.Shape)slideShape.CloneNode(true);
        SlideLayoutPart layoutPart = slide.SlidePart.SlideLayoutPart!;
        layoutPart.SlideLayout.CommonSlideData!.ShapeTree!.Append(layoutShape);
        layoutPart.SlideLayout.Save();
        layoutText.Remove();

        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointVisualReview
        });

        Assert.Contains("Layout footer", html, StringComparison.Ordinal);
        Assert.Contains("left:24pt;top:200pt;width:200pt;height:24pt", html, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerPointHtml_VisualReviewPreservesPositionedTextBoxLineBreaks() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        slide.AddTextBoxPoints("Agenda\n  Owner", 72, 96, 240, 60);

        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointVisualReview
        });

        Assert.Contains("white-space:pre-wrap", html, StringComparison.Ordinal);
        Assert.Contains("Agenda\n  Owner", html, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerPointHtml_VisualReviewPreservesShapeFlipTransforms() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointTextBox textBox = slide.AddTextBoxPoints("Flipped", 72, 96, 240, 60);
        textBox.Rotation = 12.5D;
        textBox.HorizontalFlip = true;
        textBox.VerticalFlip = true;

        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointVisualReview
        });

        Assert.Contains("transform:rotate(12.5deg) scaleX(-1) scaleY(-1);", html, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerPointHtml_CapabilityGalleryWritesSharedManifestForRichPresentation() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        slide.AddTitle("Gallery Roadmap");
        slide.AddTextBoxPoints("Positioned proof", 72, 96, 240, 60);
        slide.Notes.Text = "Gallery notes";
        using (var image = new MemoryStream(OnePixelPng)) {
            PowerPointPicture picture = slide.AddPicturePoints(image, OfficeIMO.PowerPoint.ImagePartType.Png, 72, 180, 72, 72);
            picture.AltText = "Gallery picture";
        }

        PowerPointTable table = slide.AddTablePoints(2, 2, 72, 280, 220, 80);
        table.GetCell(0, 0).Text = "Milestone";
        table.GetCell(0, 1).Text = "State";
        table.GetCell(1, 0).Text = "HTML";
        table.GetCell(1, 1).Text = "Rich proof";

        OfficeChartData chartData = new(
            new[] { "Q1", "Q2", "Q3" },
            new[] { new OfficeChartSeries("Actual", new[] { 8D, 13D, 21D }) });
        slide.AddChartPoints(OfficeChartKind.ColumnClustered, chartData, 180, 180, 260, 150).SetTitle("Gallery Pipeline");

        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.HtmlOfficeAdapters", Guid.NewGuid().ToString("N"));
        HtmlCapabilityGalleryManifest manifest = presentation.SaveHtmlCapabilityGallery(directory, new PowerPointHtmlCapabilityGalleryOptions {
            ScenarioId = "powerpoint-gallery-rich",
            Title = "PowerPoint Gallery Rich"
        });

        Assert.Equal("powerpoint-gallery-rich", manifest.Result.Scenario.Id);
        Assert.Equal(HtmlConversionProfile.PositionedReview, manifest.Profile);
        Assert.Equal(new[] {
            OfficeHtmlConversionProfile.PowerPointSemanticSlides,
            OfficeHtmlConversionProfile.PowerPointVisualReview
        }, manifest.OfficeProfiles);
        Assert.Equal(2, manifest.Result.Artifacts.Count);
        Assert.Contains(manifest.Expectations, expectation => expectation.Feature == "text boxes" && expectation.Outcome == HtmlCapabilityGalleryExpectationOutcome.Preserved);
        Assert.Contains(manifest.Expectations, expectation => expectation.Feature == "tables" && expectation.Outcome == HtmlCapabilityGalleryExpectationOutcome.Preserved);
        Assert.Contains(manifest.Expectations, expectation => expectation.Feature == "pictures" && expectation.Outcome == HtmlCapabilityGalleryExpectationOutcome.VisualProof);
        Assert.Contains(manifest.Expectations, expectation => expectation.Feature == "charts" && expectation.Outcome == HtmlCapabilityGalleryExpectationOutcome.Preserved);
        Assert.Contains(manifest.Result.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == "PowerPointChartVisualReviewRendered");
        Assert.Contains(manifest.Result.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == "PowerPointChartSemanticDataPreserved");
        Assert.DoesNotContain(manifest.Result.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == "PowerPointChartVisualPlaceholder");

        string semanticPath = Path.Combine(directory, "powerpoint-gallery-rich.semantic.html");
        string visualPath = Path.Combine(directory, "powerpoint-gallery-rich.visual.html");
        string manifestJsonPath = Path.Combine(directory, "powerpoint-gallery-rich.manifest.json");
        Assert.True(File.Exists(semanticPath));
        Assert.True(File.Exists(visualPath));
        Assert.True(File.Exists(manifestJsonPath));
        string semanticHtml = File.ReadAllText(semanticPath);
        Assert.Contains("Gallery notes", semanticHtml, StringComparison.Ordinal);
        Assert.Contains("Milestone", semanticHtml, StringComparison.Ordinal);
        Assert.Contains("officeimo-chart-data", semanticHtml, StringComparison.Ordinal);
        Assert.Contains("<th>Q2</th>", semanticHtml, StringComparison.Ordinal);
        Assert.Contains("<td>13</td>", semanticHtml, StringComparison.Ordinal);
        Assert.Contains("Rich proof", File.ReadAllText(visualPath), StringComparison.Ordinal);
        string visualHtml = File.ReadAllText(visualPath);
        Assert.Contains("officeimo-chart-rendered", visualHtml, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-visual-owner=\"OfficeIMO.Drawing\"", visualHtml, StringComparison.Ordinal);
        Assert.Contains("<svg", visualHtml, StringComparison.OrdinalIgnoreCase);

        string manifestJson = File.ReadAllText(manifestJsonPath);
        using JsonDocument json = JsonDocument.Parse(manifestJson);
        JsonElement root = json.RootElement;
        Assert.Equal("officeimo.html.capability-gallery", root.GetProperty("schemaId").GetString());
        Assert.Equal("powerpoint-gallery-rich", root.GetProperty("scenario").GetProperty("id").GetString());
        Assert.Equal("PositionedReview", root.GetProperty("profile").GetProperty("id").GetString());
        JsonElement officeProfiles = root.GetProperty("officeProfiles");
        Assert.Equal(2, officeProfiles.GetArrayLength());
        Assert.Equal("PowerPointSemanticSlides", officeProfiles[0].GetProperty("id").GetString());
        Assert.Equal("PowerPointVisualReview", officeProfiles[1].GetProperty("id").GetString());
        Assert.Equal("OfficeIMO.Drawing", officeProfiles[1].GetProperty("visualPrimitiveOwner").GetString());
        Assert.Equal(2, root.GetProperty("artifacts").GetArrayLength());
        Assert.Contains("tables", manifestJson, StringComparison.Ordinal);
        Assert.Contains("PowerPointChartSemanticDataPreserved", manifestJson, StringComparison.Ordinal);
        Assert.Contains("PowerPointChartVisualReviewRendered", manifestJson, StringComparison.Ordinal);
        Assert.DoesNotContain("PowerPointChartVisualPlaceholder", manifestJson, StringComparison.Ordinal);
        Assert.Contains("positioned review", manifestJson, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerPointHtml_LoadsSemanticRichPresentationBackToNativePresentation() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        slide.AddTitle("Roundtrip Roadmap");
        slide.AddTextBoxPoints("HTML end to end", 72, 96, 240, 60);
        slide.Notes.Text = "Presenter reminder";
        using (var image = new MemoryStream(OnePixelPng)) {
            PowerPointPicture picture = slide.AddPicturePoints(image, OfficeIMO.PowerPoint.ImagePartType.Png, 72, 180, 72, 72);
            picture.Name = "Architecture badge";
            picture.AltText = "Reusable renderer badge";
        }

        PowerPointTable table = slide.AddTablePoints(2, 2, 72, 280, 220, 80);
        table.GetCell(0, 0).Text = "Milestone";
        table.GetCell(0, 1).Text = "State";
        table.GetCell(1, 0).Text = "HTML";
        table.GetCell(1, 1).Text = "Rich proof";

        OfficeChartData chartData = new(
            new[] { "Q1", "Q2", "Q3" },
            new[] { new OfficeChartSeries("Actual", new[] { 10D, 18D, 24D }) });
        slide.AddChartPoints(OfficeChartKind.ColumnClustered, chartData, 180, 130, 260, 150).SetTitle("Pipeline");

        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides
        });

        Assert.Contains("Position: 72pt, 180pt", html, StringComparison.Ordinal);
        Assert.Contains("Position: 180pt, 130pt", html, StringComparison.Ordinal);
        Assert.Contains("Size: 260pt x 150pt", html, StringComparison.Ordinal);
        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using PowerPointPresentation imported = result.Value;
        PowerPointSlide importedSlide = imported.Slides[0];

        Assert.Equal(1, result.Slides);
        Assert.True(result.TextBoxes >= 2);
        Assert.Equal(1, result.Tables);
        Assert.Equal(1, result.Pictures);
        Assert.Equal(1, result.Charts);
        Assert.Equal(1, result.Notes);
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Message.Contains("placeholder values", StringComparison.Ordinal));
        Assert.Contains(importedSlide.TextBoxes, textBox => textBox.Text.Contains("Roundtrip Roadmap", StringComparison.Ordinal));
        Assert.Contains(importedSlide.TextBoxes, textBox => textBox.Text.Contains("HTML end to end", StringComparison.Ordinal));
        Assert.Contains(importedSlide.Tables, importedTable => importedTable.GetCell(1, 1).Text == "Rich proof");
        PowerPointPicture importedPicture = Assert.Single(importedSlide.Pictures);
        Assert.Equal("Reusable renderer badge", importedPicture.AltText);
        Assert.Equal(72D, importedPicture.LeftPoints, 3);
        Assert.Equal(180D, importedPicture.TopPoints, 3);
        PowerPointChart importedChart = Assert.Single(importedSlide.Charts);
        Assert.Equal(180D, importedChart.LeftPoints, 3);
        Assert.Equal(130D, importedChart.TopPoints, 3);
        Assert.Equal(260D, importedChart.WidthPoints, 3);
        Assert.Equal(150D, importedChart.HeightPoints, 3);
        Assert.True(importedChart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(new[] { "Q1", "Q2", "Q3" }, snapshot.Data.Categories);
        OfficeChartSeries importedSeries = Assert.Single(snapshot.Data.Series);
        Assert.Equal("Actual", importedSeries.Name);
        Assert.Equal(new[] { 10D, 18D, 24D }, importedSeries.Values);
        Assert.Contains("Presenter reminder", importedSlide.Notes.Text, StringComparison.Ordinal);

        string roundTripHtml = imported.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides
        });

        Assert.Contains("data-officeimo-profile=\"PowerPointSemanticSlides\"", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("Roundtrip Roadmap", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("Rich proof", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("Reusable renderer badge", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("Pipeline", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("officeimo-chart-data", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("<th>Q2</th>", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("<td>18</td>", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("Presenter reminder", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("data:image/png;base64", roundTripHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerPointHtml_RoundTripsOffSlidePicturePosition() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        using (var image = new MemoryStream(OnePixelPng)) {
            slide.AddPicturePoints(image, OfficeIMO.PowerPoint.ImagePartType.Png, -18, -12, 72, 72);
        }

        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides
        });

        Assert.Contains("Position: -18pt, -12pt", html, StringComparison.Ordinal);
        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using PowerPointPresentation imported = result.Value;

        PowerPointPicture picture = Assert.Single(imported.Slides[0].Pictures);
        Assert.Equal(-18D, picture.LeftPoints, 3);
        Assert.Equal(-12D, picture.TopPoints, 3);
    }

    [Fact]
    public void PowerPointHtml_RoundTripsPictureTransforms() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        using (var image = new MemoryStream(OnePixelPng)) {
            PowerPointPicture picture = slide.AddPicturePoints(image, OfficeIMO.PowerPoint.ImagePartType.Png, 80, 90, 120, 72);
            picture.Name = "Transformed picture";
            picture.AltText = "Transformed alt";
            picture.Rotation = 23.5D;
            picture.HorizontalFlip = true;
            picture.VerticalFlip = true;
            picture.Crop(10D, 20D, 5D, 15D);
        }

        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides
        });

        Assert.Contains("data-officeimo-rotation=\"23.5\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-flip-horizontal=\"true\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-flip-vertical=\"true\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-crop-left=\"0.10000000000000001\"", html, StringComparison.Ordinal);
        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using PowerPointPresentation imported = result.Value;

        PowerPointPicture importedPicture = Assert.Single(imported.Slides[0].Pictures);
        Assert.Equal("Transformed alt", importedPicture.AltText);
        Assert.Equal(80D, importedPicture.LeftPoints, 3);
        Assert.Equal(90D, importedPicture.TopPoints, 3);
        Assert.Equal(120D, importedPicture.WidthPoints, 3);
        Assert.Equal(72D, importedPicture.HeightPoints, 3);
        Assert.Equal(23.5D, importedPicture.Rotation!.Value, 3);
        Assert.True(importedPicture.HorizontalFlip);
        Assert.True(importedPicture.VerticalFlip);
        Assert.Equal(0.1D, importedPicture.CropLeftRatio, 3);
        Assert.Equal(0.2D, importedPicture.CropTopRatio, 3);
        Assert.Equal(0.05D, importedPicture.CropRightRatio, 3);
        Assert.Equal(0.15D, importedPicture.CropBottomRatio, 3);
    }

    [Fact]
    public void PowerPointHtml_RejectsCropPairsThatConsumeTheWholePicture() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        using (var image = new MemoryStream(OnePixelPng)) {
            slide.AddPicturePoints(image, OfficeIMO.PowerPoint.ImagePartType.Png, 40, 40, 80, 60)
                .Crop(10D, 0D, 10D, 0D);
        }
        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides
        });
        html = System.Text.RegularExpressions.Regex.Replace(
            html, "data-officeimo-crop-left=\"[^\"]*\"", "data-officeimo-crop-left=\"0.75\"");
        html = System.Text.RegularExpressions.Regex.Replace(
            html, "data-officeimo-crop-right=\"[^\"]*\"", "data-officeimo-crop-right=\"0.75\"");

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using PowerPointPresentation imported = result.Value;
        PowerPointPicture picture = Assert.Single(imported.Slides[0].Pictures);

        Assert.Equal(0D, picture.CropLeftRatio);
        Assert.Equal(0D, picture.CropRightRatio);
        Assert.Contains(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticValueInvalid);
    }

    [Fact]
    public void PowerPointHtml_RoundTripsChartTransformsAndMixedDrawingOrder() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        OfficeChartData chartData = new(
            new[] { "Q1", "Q2" },
            new[] { new OfficeChartSeries("Actual", new[] { 10D, 18D }) });
        PowerPointChart chart = slide.AddChartPoints(OfficeChartKind.ColumnClustered, chartData, 120, 90, 240, 140);
        chart.SetTitle("Transform chart");
        chart.Rotation = 18.75D;
        chart.HorizontalFlip = true;
        chart.VerticalFlip = true;
        using (var image = new MemoryStream(OnePixelPng)) {
            PowerPointPicture picture = slide.AddPicturePoints(image, OfficeIMO.PowerPoint.ImagePartType.Png, 96, 100, 80, 64);
            picture.Name = "Top picture";
        }

        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides
        });

        Assert.Contains("data-officeimo-layer-kind=\"chart\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-layer-kind=\"picture\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rotation=\"18.75\"", html, StringComparison.Ordinal);
        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using PowerPointPresentation imported = result.Value;
        PowerPointSlide importedSlide = imported.Slides[0];

        Assert.Equal(1, result.Charts);
        Assert.Equal(1, result.Pictures);
        Assert.Empty(result.Report.Diagnostics);
        PowerPointChart importedChart = Assert.Single(importedSlide.Charts);
        PowerPointPicture importedPicture = Assert.Single(importedSlide.Pictures);
        Assert.Equal(18.75D, importedChart.Rotation!.Value, 3);
        Assert.True(importedChart.HorizontalFlip);
        Assert.True(importedChart.VerticalFlip);
        Assert.True(importedChart.DrawingOrder < importedPicture.DrawingOrder);
    }

    [Fact]
    public void PowerPointHtml_VisualReviewStretchesPositionedPicturesToAuthoredShape() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        using (var image = new MemoryStream(OnePixelPng)) {
            slide.AddPicturePoints(image, OfficeIMO.PowerPoint.ImagePartType.Png, 40, 50, 180, 60);
        }

        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointVisualReview
        });

        Assert.Contains(".officeimo-shape-picture img{width:100%;height:100%;object-fit:fill;display:block;}", html, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerPointHtml_LoadCreatesDistinctSlidesForEachSemanticSection() {
        string html = """
            <main>
              <section class="officeimo-slide"><p>First slide</p></section>
              <section class="officeimo-slide"><p>Second slide</p></section>
            </main>
            """;

        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using PowerPointPresentation imported = result.Value;

        Assert.Equal(2, result.Slides);
        Assert.Equal(2, imported.Slides.Count);
        Assert.Contains(imported.Slides[0].TextBoxes, textBox => textBox.Text.Contains("First slide", StringComparison.Ordinal));
        Assert.Contains(imported.Slides[1].TextBoxes, textBox => textBox.Text.Contains("Second slide", StringComparison.Ordinal));
    }

    [Fact]
    public void PowerPointHtml_RoundTripsHiddenSlideStateWhenIncluded() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        slide.AddTextBox("Hidden briefing");
        slide.Hidden = true;

        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides,
            IncludeHiddenSlides = true
        });

        Assert.Contains("data-officeimo-hidden=\"true\"", html, StringComparison.Ordinal);

        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using PowerPointPresentation imported = result.Value;
        PowerPointSlide importedSlide = Assert.Single(imported.Slides);
        Assert.True(importedSlide.Hidden);
        Assert.Contains(importedSlide.TextBoxes, textBox => textBox.Text.Contains("Hidden briefing", StringComparison.Ordinal));
    }

    [Fact]
    public void PowerPointHtml_LoadPreservesSlideTextWhitespace() {
        string html = """
            <main>
              <section class="officeimo-slide">
                <p>Line 1&#10;Line 2  with  spacing</p>
                <table>
                  <tr><td>A  B</td><td>First&#10;Second</td></tr>
                </table>
              </section>
            </main>
            """;

        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using PowerPointPresentation imported = result.Value;
        PowerPointSlide slide = imported.Slides[0];
        PowerPointTextBox textBox = Assert.Single(slide.TextBoxes);
        PowerPointTable table = Assert.Single(slide.Tables);

        Assert.Equal("Line 1\nLine 2  with  spacing", textBox.Text);
        Assert.Equal("A  B", table.GetCell(0, 0).Text);
        Assert.Equal("First\nSecond", table.GetCell(0, 1).Text);
    }

    [Fact]
    public void PowerPointHtml_LoadPreservesChartKindFromSemanticInventory() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        OfficeChartData chartData = new(
            new[] { "Q1", "Q2" },
            new[] { new OfficeChartSeries("Actual", new[] { 10D, 12D }) });
        slide.AddChartPoints(OfficeChartKind.Line, chartData, 72, 96, 240, 140).SetTitle("Trend");

        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides
        });

        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using PowerPointPresentation imported = result.Value;
        PowerPointChart importedChart = Assert.Single(imported.Slides[0].Charts);

        Assert.True(importedChart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(OfficeChartKind.Line, snapshot.ChartKind);
    }

    [Fact]
    public void PowerPointHtml_LoadPreservesBlankChartCategories() {
        string html = """
            <main>
              <section class="officeimo-slide">
                <section class="officeimo-feature officeimo-charts">
                  <ul class="officeimo-feature-list">
                    <li class="officeimo-feature-item">
                      <span class="officeimo-feature-label">Blank category</span>
                      <div class="officeimo-feature-meta">Type: ClusteredColumn; Series: 1; Categories: 3</div>
                      <table class="officeimo-chart-data">
                        <thead><tr><th></th><th>Q1</th><th></th><th>Q3</th></tr></thead>
                        <tbody><tr><th>Actual</th><td>10</td><td>12</td><td>14</td></tr></tbody>
                      </table>
                    </li>
                  </ul>
                </section>
              </section>
            </main>
            """;

        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using PowerPointPresentation imported = result.Value;
        PowerPointChart importedChart = Assert.Single(imported.Slides[0].Charts);

        Assert.True(importedChart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(new[] { "Q1", string.Empty, "Q3" }, snapshot.Data.Categories);
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Message.Contains("placeholder values", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("StackedLine", OfficeChartKind.LineStacked)]
    [InlineData("StackedLine100", OfficeChartKind.LineStacked100)]
    public void PowerPointHtml_LoadPreservesStackedLineChartKinds(string chartKind, OfficeChartKind expectedKind) {
        string html = $$"""
            <main>
              <section class="officeimo-slide">
                <section class="officeimo-feature officeimo-charts">
                  <ul class="officeimo-feature-list">
                    <li class="officeimo-feature-item">
                      <span class="officeimo-feature-label">Stacked trend</span>
                      <div class="officeimo-feature-meta">Type: {{chartKind}}; Series: 2; Categories: 2</div>
                      <table class="officeimo-chart-data">
                        <thead><tr><th></th><th>Q1</th><th>Q2</th></tr></thead>
                        <tbody>
                          <tr><th>Actual</th><td>10</td><td>12</td></tr>
                          <tr><th>Plan</th><td>8</td><td>11</td></tr>
                        </tbody>
                      </table>
                    </li>
                  </ul>
                </section>
              </section>
            </main>
            """;

        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using PowerPointPresentation imported = result.Value;
        PowerPointChart importedChart = Assert.Single(imported.Slides[0].Charts);

        Assert.True(importedChart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(expectedKind, snapshot.ChartKind);
        Assert.Equal(2, snapshot.Data.Series.Count);
    }

    [Theory]
    [InlineData("ClusteredBar")]
    [InlineData("StackedColumn")]
    [InlineData("Area")]
    [InlineData("Radar")]
    public void PowerPointHtml_LoadDegradesExportedChartKindsInsteadOfDroppingChart(string chartKind) {
        string html = $$"""
            <main>
              <section class="officeimo-slide">
                <section class="officeimo-feature officeimo-charts">
                  <ul class="officeimo-feature-list">
                    <li class="officeimo-feature-item">
                      <span class="officeimo-feature-label">Imported chart</span>
                      <div class="officeimo-feature-meta">Type: {{chartKind}}; Series: 1; Categories: 2</div>
                      <table class="officeimo-chart-data">
                        <thead><tr><th></th><th>Q1</th><th>Q2</th></tr></thead>
                        <tbody><tr><th>Actual</th><td>10</td><td>12</td></tr></tbody>
                      </table>
                    </li>
                  </ul>
                </section>
              </section>
            </main>
            """;

        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using PowerPointPresentation imported = result.Value;
        PowerPointChart importedChart = Assert.Single(imported.Slides[0].Charts);

        Assert.Equal(1, result.Charts);
        Assert.True(importedChart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(OfficeChartKind.ColumnClustered, snapshot.ChartKind);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Message.Contains("used chart kind '" + chartKind + "' and was imported as a clustered column fallback", StringComparison.Ordinal));
    }

    [Fact]
    public void PowerPointHtml_LoadPreservesHeadingLikePresenterNotes() {
        string html = """
            <main>
              <section class="officeimo-slide">
                <p>Notes slide</p>
                <pre class="officeimo-source-markdown"># Notes slide

            ## Body section
            ### Notes
            This is regular slide content, not presenter notes.

            ### Notes
            # Follow up
            Keep this line</pre>
              </section>
            </main>
            """;

        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using PowerPointPresentation imported = result.Value;

        Assert.Equal(1, result.Notes);
        Assert.DoesNotContain("This is regular slide content", imported.Slides[0].Notes.Text, StringComparison.Ordinal);
        Assert.Contains("# Follow up", imported.Slides[0].Notes.Text, StringComparison.Ordinal);
        Assert.Contains("Keep this line", imported.Slides[0].Notes.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerPointHtml_LoadPreservesPresenterNoteParagraphBreaks() {
        string html = """
            <main>
              <section class="officeimo-slide">
                <p>Notes slide</p>
                <pre class="officeimo-source-markdown"># Notes slide

            ### Notes
            First line

                Indented line
            Third line</pre>
              </section>
            </main>
            """;

        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using PowerPointPresentation imported = result.Value;

        string expected = string.Join(Environment.NewLine, "First line", string.Empty, "    Indented line", "Third line");
        Assert.Equal(1, result.Notes);
        Assert.Equal(expected, imported.Slides[0].Notes.Text);
    }

    [Fact]
    public void PowerPointHtml_LoadPreservesNonPngImageMediaTypes() {
        string image = Convert.ToBase64String(OnePixelPng);
        string html = $"""
            <main>
              <section class="officeimo-slide">
                <section class="officeimo-feature officeimo-images">
                  <ul class="officeimo-feature-list">
                    <li class="officeimo-feature-item">
                      <span class="officeimo-feature-label">Tiff marker</span>
                      <div class="officeimo-feature-meta">Size: 24pt x 24pt; Type: image/tif</div>
                      <img src="data:image/tif;base64,{image}" alt="Tiff marker" />
                    </li>
                  </ul>
                </section>
              </section>
            </main>
            """;

        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using PowerPointPresentation imported = result.Value;
        PowerPointPicture picture = Assert.Single(imported.Slides[0].Pictures);

        Assert.Equal(1, result.Pictures);
        Assert.Equal("image/tiff", picture.ContentType);
    }

    [Fact]
    public void OfficeHtmlDocumentShell_UsesSharedThemes() {
        string html = OfficeHtmlDocumentShell.WrapBody(
            "<main class=\"officeimo-document\"><p>Styled</p></main>",
            new OfficeHtmlDocumentOptions {
                Title = "Theme",
                Theme = OfficeVisualThemeKind.TechnicalDocument
            });

        Assert.Contains("<title>Theme</title>", html, StringComparison.Ordinal);
        Assert.Contains("--officeimo-accent:#047857", html, StringComparison.Ordinal);
        Assert.Contains("Styled", html, StringComparison.Ordinal);
    }
}
