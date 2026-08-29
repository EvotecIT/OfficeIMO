using OfficeIMO.Excel.Legacy;
using OfficeIMO.Excel.Html;
using OfficeIMO.Excel.Csv;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Excel;

namespace OfficeIMO.LegacyImport.Tests;

public sealed class LegacySpreadsheetImportTests {
    public static IEnumerable<object[]> Families() {
        yield return new object[] { LegacyFixtureFactory.Wk(), "archive.wk1", LegacySpreadsheetFormat.Lotus123 };
        yield return new object[] { LegacyFixtureFactory.Wk(0x20, 0x51), "archive.wq1", LegacySpreadsheetFormat.QuattroPro };
        yield return new object[] { LegacyFixtureFactory.Multiplan(), "archive.mp", LegacySpreadsheetFormat.Multiplan };
        yield return new object[] { LegacyFixtureFactory.Wk(includeFormulaAndChart: false), "archive.wks", LegacySpreadsheetFormat.MicrosoftWorks };
        yield return new object[] { LegacyFixtureFactory.CompoundSheet(), "archive.xlr", LegacySpreadsheetFormat.MicrosoftWorks };
    }

    [Theory]
    [MemberData(nameof(Families))]
    public void DetectsAndImportsEveryBoundedFamily(byte[] source, string sourceName, LegacySpreadsheetFormat expected) {
        using LegacySpreadsheetImportResult result = LegacySpreadsheetImporter.Import(source, new LegacySpreadsheetImportOptions { SourceName = sourceName });
        Assert.Equal(expected, result.Detection.Format);
        Assert.True(result.Report.RecoveredItemCount > 0);
        Assert.NotEmpty(result.Document.Sheets);
    }

    [Fact]
    public void WkRecordsRecoverCachedValuesAlignmentAndChartMetadata() {
        using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import(LegacyFixtureFactory.Wk(), new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" });
        Assert.Equal(OfficeLegacyImportQuality.Structured, imported.Report.Quality);
        Assert.Single(imported.Charts);
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "LEGACY_FORMULA_CACHED_VALUE");
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "LEGACY_CHART_METADATA_ONLY");
        var sheet = Assert.Single(imported.Document.CreateInspectionSnapshot().Worksheets);
        Assert.Equal("Name", Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1).Value);
        Assert.Equal("left", Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1).Style?.HorizontalAlignment);
        Assert.Equal(42, Convert.ToInt32(Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2).Value));
        Assert.Equal(84d, Convert.ToDouble(Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 3).Value));
    }

    [Fact]
    public void WkParserStopsAtEofAndBoundsMetadataText() {
        byte[] trailingCell = { 0x0D, 0x00, 0x07, 0x00, 0x00, 0x03, 0x00, 0x00, 0x00, 0x63, 0x00 };
        byte[] source = LegacyFixtureFactory.Wk().Concat(trailingCell).ToArray();
        using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import(source, new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" });
        Assert.Equal(3, imported.Report.RecoveredItemCount);

        Assert.Throws<InvalidDataException>(() => LegacySpreadsheetImporter.Import(LegacyFixtureFactory.Wk(), new LegacySpreadsheetImportOptions {
            SourceName = "archive.wk1",
            Limits = new OfficeLegacyImportLimits { MaxTextCharacters = 3 }
        }));

        using LegacySpreadsheetImportResult formatted = LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.Wk(cellFormat: 1),
            new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" });
        Assert.Contains(formatted.Report.Findings, finding => finding.Code == "LEGACY_CELL_FORMAT_PARTIAL");
    }

    [Fact]
    public void WeakExtensionsAndUninspectableCompoundSecurityDoNotPassSilently() {
        Assert.Throws<InvalidDataException>(() => LegacySpreadsheetImporter.Import(
            Encoding.ASCII.GetBytes("renamed,plain,text"),
            new LegacySpreadsheetImportOptions { SourceName = "renamed.wk1" }));

        using LegacySpreadsheetImportResult compound = LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.CompoundSheet(),
            new LegacySpreadsheetImportOptions { SourceName = "archive.xlr" });
        Assert.Contains(compound.Report.Findings, finding => finding.Code == "LEGACY_COMPOUND_INVENTORY_INCOMPLETE");
        Assert.True(compound.Report.HasLoss);
    }

    [Fact]
    public void ImportedWorkbookUsesEverySupportedModernOutputOwner() {
        using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import(LegacyFixtureFactory.Wk(), new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" });
        using var xlsx = new MemoryStream();
        imported.Document.Save(xlsx);
        Assert.True(xlsx.Length > 100);
        Assert.Contains("Name", imported.Document.ToHtml());
        Assert.Contains("Name", imported.Document.Sheets[0].ToCsv());
        Assert.StartsWith("%PDF", Encoding.ASCII.GetString(imported.Document.ToPdf(), 0, 4));
        using var ods = new MemoryStream();
        imported.Document.ToOpenDocument().Save(ods);
        Assert.True(ods.Length > 100);
    }

    [Fact]
    public void ReaderHandlerProjectsLegacyWarningsAndContent() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddLegacySpreadsheetHandler().Build();
        using var stream = new MemoryStream(LegacyFixtureFactory.Wk());
        OfficeDocumentReadResult result = reader.ReadDocument(stream, "archive.wk1");
        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("Name", StringComparison.Ordinal));
        Assert.Contains(OfficeDocumentReaderBuilderExcelExtensions.LegacyHandlerId, result.CapabilitiesUsed);
        Assert.Contains(result.Chunks.SelectMany(chunk => chunk.Warnings ?? Array.Empty<string>()), warning => warning.Contains("Legacy import quality", StringComparison.Ordinal));
    }

    [Fact]
    public void ImportHonorsCancellation() {
        Assert.Throws<OperationCanceledException>(() => LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.Wk(),
            new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" },
            new CancellationToken(canceled: true)));
    }

    [Fact]
    public void ReaderRegistrationCapturesLegacyImportOptions() {
        var options = new LegacySpreadsheetImportOptions {
            FormatHint = LegacySpreadsheetFormat.Lotus123,
            Limits = new OfficeLegacyImportLimits { MaxInputBytes = 1024 }
        };
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddLegacySpreadsheetHandler(options).Build();
        options.FormatHint = LegacySpreadsheetFormat.Multiplan;
        options.Limits.MaxInputBytes = 1;

        using var stream = new MemoryStream(LegacyFixtureFactory.Wk());
        OfficeDocumentReadResult result = reader.ReadDocument(stream, "archive.wk1");
        Assert.Contains(result.Chunks.SelectMany(chunk => chunk.Warnings ?? Array.Empty<string>()),
            warning => warning.Contains("lotus-1-2-3", StringComparison.Ordinal));
    }
}
