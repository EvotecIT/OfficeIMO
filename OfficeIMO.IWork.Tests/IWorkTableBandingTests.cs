using System.Text.Json;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Native_Numbers_banding_and_region_intersections_match_saved_Apple_exports() {
        using var manifest = JsonDocument.Parse(File.ReadAllText(CorpusFixture("native-exports/numbers-banding-v14.5.json")));
        var root = manifest.RootElement;
        string source = CorpusFixture(root.GetProperty("sourceFixture").GetString()!);
        Assert.Equal(root.GetProperty("sourceSha256").GetString(), HashFile(source));
        foreach (var artifact in root.GetProperty("artifacts").EnumerateArray())
            Assert.Equal(artifact.GetProperty("sha256").GetString(), HashFile(CorpusFixture("native-exports/" + artifact.GetProperty("path").GetString())));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(source,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.False(result.IsVisualFallback);
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_FILL_DEFAULTS_UNSUPPORTED");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
        foreach (var expected in root.GetProperty("tables").EnumerateArray()) {
            string name = expected.GetProperty("sheet").GetString()!;
            var sheet = Assert.Single(result.Projection.Sheets, sheet => sheet.Name == name);
            var table = Assert.Single(sheet.Tables);
            var destination = reopened.Sheets[result.Projection.Sheets.ToList().IndexOf(sheet)];
            Assert.Equal(expected.GetProperty("modelIdentifier").GetUInt64(), table.ModelRecord!.Identifier);
            Assert.Equal(expected.GetProperty("headerRows").GetInt32(), table.HeaderRowCount);
            Assert.Equal(expected.GetProperty("bandedBodyRgb").GetString(), table.FillStyles.BandedBody!.Color!.RgbHex);
            Assert.Null(table.GetCell(8, 3)); // Unstored footer cells still carry the region fill.
            foreach (var cell in expected.GetProperty("cells").EnumerateArray()) {
                int row = cell.GetProperty("row").GetInt32(), column = cell.GetProperty("column").GetInt32();
                string? rgb = cell.GetProperty("rgb").GetString();
                Assert.Equal(rgb, table.GetFill(row, column)?.Color?.RgbHex);
                Assert.Equal(rgb == null ? null : "FF" + rgb, destination.GetCellStyle(row, column).FillColorArgb);
            }
        }
    }

    [Fact]
    public void Banded_body_fill_alone_is_bounded_by_the_destination_style_budget() {
        using var package = RoleFillPackage(IWorkDocumentKind.Numbers, rows: 100_001, banded: true, roleDefaults: false);
        using var result = IWorkSourceDocument.Open(package).ToExcelDocumentResult(
            new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Message.Contains("styled-cell budget", StringComparison.Ordinal));
    }

    [Fact]
    public void Selected_no_fill_and_unresolved_fill_suppress_an_active_band_without_materializing_the_grid() {
        var band = new IWorkCellFill(new IWorkColor(128, 128, 128, 255));
        var none = new IWorkCellFill(null);
        var table = new IWorkTable("Selected", 4, 2, new[] {
            new IWorkTableCell(2, 1, IWorkCellKind.Empty, null, fill: none),
            new IWorkTableCell(2, 2, IWorkCellKind.Empty, null, hasUnresolvedFill: true)
        }, fillStyles: new IWorkTableFillStyles(null, null, null, null, band));
        Assert.Same(none, table.GetFill(2, 1));
        Assert.Null(table.GetFill(2, 2));
        Assert.Same(band, table.GetFill(4, 1));
        Assert.Equal(2, table.Cells.Count);
    }

    [Fact]
    public void Disabled_banding_ignores_an_inactive_unsupported_band_fill() {
        using var package = RoleFillPackage(IWorkDocumentKind.Pages, "inactive-band-gradient");
        var projection = IWorkSourceDocument.Open(package).ReadPages();
        var table = Assert.Single(projection.Tables);
        Assert.Null(table.FillStyles.BandedBody);
        Assert.Equal("0000FF", table.GetFill(3, 2)!.Color!.RgbHex);
        Assert.DoesNotContain(projection.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_FILL_DEFAULTS_UNSUPPORTED");
        Assert.DoesNotContain(projection.SourceDeclarationIssues, issue => issue.FieldPath == "11/2");
    }

    [Fact]
    public void Active_banding_inherits_the_parent_fill_when_the_child_only_enables_banding() {
        using var package = RoleFillPackage(IWorkDocumentKind.Numbers, "inherited-band");
        var projection = IWorkSourceDocument.Open(package).ReadNumbers();
        var table = Assert.Single(Assert.Single(projection.Sheets).Tables);
        Assert.Equal("808080", table.FillStyles.BandedBody!.Color!.RgbHex);
        Assert.Equal("808080", table.GetFill(3, 2)!.Color!.RgbHex);
        Assert.Equal("0000FF", table.GetFill(2, 3)!.Color!.RgbHex);
        Assert.DoesNotContain(projection.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_FILL_DEFAULTS_UNSUPPORTED");
    }

    private static string HashFile(string path) => Convert.ToHexString(
        System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
}
