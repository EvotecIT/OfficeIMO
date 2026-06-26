using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
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
        ExcelSheet sheet = workbook.AddWorkSheet("Sales");
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
        sheet.AddImage(7, 1, OnePixelPng, widthPixels: 48, heightPixels: 48, name: "Status Logo", altText: "Inline status marker");

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables,
            Theme = OfficeHtmlDocumentThemeKind.Report
        });

        Assert.Contains("data-officeimo-profile=\"ExcelSemanticTables\"", html, StringComparison.Ordinal);
        Assert.Contains("<th>Region</th>", html, StringComparison.Ordinal);
        Assert.Contains("<td>North</td>", html, StringComparison.Ordinal);
        Assert.Contains("123.45", html, StringComparison.Ordinal);
        Assert.Contains("--officeimo-accent", html, StringComparison.Ordinal);
        Assert.Contains("officeimo-formulas", html, StringComparison.Ordinal);
        Assert.Contains("SUM(B2:B4)", html, StringComparison.Ordinal);
        Assert.Contains("Reviewed with finance", html, StringComparison.Ordinal);
        Assert.Contains("Revenue Trend", html, StringComparison.Ordinal);
        Assert.Contains("officeimo-chart-data", html, StringComparison.Ordinal);
        Assert.Contains("<th>South</th>", html, StringComparison.Ordinal);
        Assert.Contains("<th>Amount</th>", html, StringComparison.Ordinal);
        Assert.Contains("Inline status marker", html, StringComparison.Ordinal);
        Assert.Contains("data:image/png;base64", html, StringComparison.Ordinal);
    }

    [Fact]
    public void ExcelHtml_ExportsVisualReviewFromSharedSvgRenderer() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorkSheet("Visual");
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
    public void ExcelHtml_CapabilityGalleryWritesSharedManifestForRichWorkbook() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorkSheet("Gallery");
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
        Assert.Contains("<th>South</th>", semanticHtml, StringComparison.Ordinal);
        Assert.Contains("<td>57</td>", semanticHtml, StringComparison.Ordinal);
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
        ExcelSheet sheet = workbook.AddWorkSheet("Roundtrip");
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
        sheet.AddImage(7, 1, OnePixelPng, widthPixels: 48, heightPixels: 48, name: "Status Logo", altText: "Inline status marker");

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables,
            Theme = OfficeHtmlDocumentThemeKind.Report
        });

        ExcelHtmlLoadResult result = html.LoadExcelFromHtmlWithResult();
        using ExcelDocument imported = result.Workbook;
        ExcelSheet importedSheet = imported.Sheets.Single(importedSheet => importedSheet.Name == "Roundtrip");

        Assert.Equal(1, result.Sheets);
        Assert.True(result.Cells >= 8);
        Assert.Equal(1, result.Formulas);
        Assert.Equal(1, result.Comments);
        Assert.Equal(1, result.Images);
        Assert.Equal(1, result.Charts);
        Assert.Empty(result.Diagnostics);
        Assert.True(importedSheet.TryGetCellText(2, 1, out string region));
        Assert.Equal("North", region);
        Assert.Contains(importedSheet.GetFormulaCells(), formula => formula.CellReference == "B5" && formula.Formula == "SUM(B2:B4)");
        Assert.Contains(importedSheet.GetComments(), comment => comment.CellReference == "B2" && comment.Text == "Reviewed with finance" && comment.Author == "OfficeIMO");
        Assert.Single(importedSheet.Images);
        ExcelChart importedChart = Assert.Single(importedSheet.Charts);
        Assert.True(importedChart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        Assert.Equal(new[] { "North", "South", "West" }, snapshot.Data.Categories);
        ExcelChartSeries importedSeries = Assert.Single(snapshot.Data.Series);
        Assert.Equal("Amount", importedSeries.Name);
        Assert.Equal(new[] { 123.45D, 98.2D, 140D }, importedSeries.Values);

        string roundTripHtml = imported.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables,
            Theme = OfficeHtmlDocumentThemeKind.Report
        });

        Assert.Contains("data-officeimo-profile=\"ExcelSemanticTables\"", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("North", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("SUM(B2:B4)", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("Reviewed with finance", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("Revenue Trend", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("officeimo-chart-data", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("<th>South</th>", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("<th>Amount</th>", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("Inline status marker", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("data:image/png;base64", roundTripHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerPointHtml_ExportsSemanticSlidesWithExtractionProof() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.Slides[0];
        slide.AddTitle("Roadmap");
        slide.AddTextBox("HTML end to end");
        slide.Notes.Text = "Presenter reminder";
        using (var image = new MemoryStream(OnePixelPng)) {
            PowerPointPicture picture = slide.AddPicturePoints(image, OfficeIMO.PowerPoint.ImagePartType.Png, 72, 140, 72, 72);
            picture.Name = "Architecture badge";
            picture.AltText = "Reusable renderer badge";
        }

        PowerPointChartData chartData = new(
            new[] { "Q1", "Q2", "Q3" },
            new[] { new PowerPointChartSeries("Actual", new[] { 10D, 18D, 24D }) });
        slide.AddChartPoints(chartData, 180, 130, 260, 150).SetTitle("Pipeline");

        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides
        });

        Assert.Contains("data-officeimo-profile=\"PowerPointSemanticSlides\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Roadmap</p>", html, StringComparison.Ordinal);
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
        PowerPointSlide slide = presentation.Slides[0];
        slide.AddTextBoxPoints("Positioned", 72, 96, 240, 60);
        using (var image = new MemoryStream(OnePixelPng)) {
            PowerPointPicture picture = slide.AddPicturePoints(image, OfficeIMO.PowerPoint.ImagePartType.Png, 72, 180, 72, 72);
            picture.AltText = "Positioned image";
        }

        PowerPointChartData chartData = new(
            new[] { "Q1", "Q2", "Q3" },
            new[] { new PowerPointChartSeries("Actual", new[] { 8D, 13D, 21D }) });
        slide.AddChartPoints(chartData, 180, 180, 260, 150).SetTitle("Positioned Pipeline");

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
    public void PowerPointHtml_CapabilityGalleryWritesSharedManifestForRichPresentation() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.Slides[0];
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

        PowerPointChartData chartData = new(
            new[] { "Q1", "Q2", "Q3" },
            new[] { new PowerPointChartSeries("Actual", new[] { 8D, 13D, 21D }) });
        slide.AddChartPoints(chartData, 180, 180, 260, 150).SetTitle("Gallery Pipeline");

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
        PowerPointSlide slide = presentation.Slides[0];
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

        PowerPointChartData chartData = new(
            new[] { "Q1", "Q2", "Q3" },
            new[] { new PowerPointChartSeries("Actual", new[] { 10D, 18D, 24D }) });
        slide.AddChartPoints(chartData, 180, 130, 260, 150).SetTitle("Pipeline");

        string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides
        });

        PowerPointHtmlLoadResult result = html.LoadPowerPointFromHtmlWithResult();
        using PowerPointPresentation imported = result.Presentation;
        PowerPointSlide importedSlide = imported.Slides[0];

        Assert.Equal(1, result.Slides);
        Assert.True(result.TextBoxes >= 2);
        Assert.Equal(1, result.Tables);
        Assert.Equal(1, result.Pictures);
        Assert.Equal(1, result.Charts);
        Assert.Equal(1, result.Notes);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Contains("placeholder values", StringComparison.Ordinal));
        Assert.Contains(importedSlide.TextBoxes, textBox => textBox.Text.Contains("Roundtrip Roadmap", StringComparison.Ordinal));
        Assert.Contains(importedSlide.TextBoxes, textBox => textBox.Text.Contains("HTML end to end", StringComparison.Ordinal));
        Assert.Contains(importedSlide.Tables, importedTable => importedTable.GetCell(1, 1).Text == "Rich proof");
        Assert.Contains(importedSlide.Pictures, picture => picture.AltText == "Reusable renderer badge");
        PowerPointChart importedChart = Assert.Single(importedSlide.Charts);
        Assert.True(importedChart.TryGetSnapshot(out PowerPointChartSnapshot? snapshot));
        Assert.Equal(new[] { "Q1", "Q2", "Q3" }, snapshot!.Data.Categories);
        PowerPointChartSeries importedSeries = Assert.Single(snapshot.Data.Series);
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
    public void OfficeHtmlDocumentShell_UsesSharedThemes() {
        string html = OfficeHtmlDocumentShell.WrapBody(
            "<main class=\"officeimo-document\"><p>Styled</p></main>",
            new OfficeHtmlDocumentOptions {
                Title = "Theme",
                Theme = OfficeHtmlDocumentThemeKind.Technical
            });

        Assert.Contains("<title>Theme</title>", html, StringComparison.Ordinal);
        Assert.Contains("--officeimo-accent:#047857", html, StringComparison.Ordinal);
        Assert.Contains("Styled", html, StringComparison.Ordinal);
    }
}
