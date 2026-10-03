using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Day_duration_formula_keeps_its_expression_and_fractional_cache_after_save() {
        byte[] cell = new byte[28]; cell[0] = 5; cell[1] = 7;
        WriteUInt32(cell, 8, (1u << 1) | (1u << 9) | (1u << 16));
        Buffer.BlockCopy(BitConverter.GetBytes(151200d), 0, cell, 12, 8);
        WriteUInt32(cell, 20, 0); WriteUInt32(cell, 24, 1);
        using var package = TableDependencyPackage(IWorkDocumentKind.Numbers,
            Message(ReferenceField(22, 13), ReferenceField(6, 14)), cellPayload: cell,
            additionalRecords: Message(
                ArchiveRecord(13, 6005, Message(VarintField(1, 2), BytesField(3, FormatEntry(DurationFormat(largest: 2, smallest: 2))))),
                ArchiveRecord(14, 6201, Message(VarintField(1, 3), BytesField(3, Message(VarintField(1, 0), BytesField(5, FormulaConstant(1.75d))))))));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        Assert.Equal(151200d, result.Projection.Sheets[0].Tables[0].Cells[0].Value);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
        Assert.Equal("1.75", reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(1.75d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Equal("2d", Assert.Single(reopened.Sheets[0].Range("A1").CreateVisualSnapshot().Cells).Text);
    }

    [Theory]
    [InlineData(8726400d, "101d")]
    [InlineData(-31017600d, "-359d")]
    [InlineData(0d, "0d")]
    [InlineData(151200d, "2d")]
    [InlineData(-151200d, "-2d")]
    public void Day_duration_preserves_numeric_values_and_shared_saved_display(double seconds, string expected) {
        using var package = DurationFormatPackage(IWorkDocumentKind.Numbers, seconds, DurationFormat(largest: 2, smallest: 2));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        Assert.False(result.IsVisualFallback);
        var cell = Assert.Single(result.Projection.Sheets[0].Tables[0].Cells);
        Assert.Equal(seconds, cell.Value);
        Assert.Equal(IWorkDurationUnit.Day, cell.NumberFormat!.DurationFormat!.LargestUnit);
        Assert.Equal(IWorkDurationUnit.Day, cell.NumberFormat.DurationFormat.SmallestUnit);
        Assert.Equal(IWorkCellUnsupportedFeatures.None, cell.UnsupportedFeatures);
        Assert.True(cell.TryGetFormattedNumber(out string display, out _)); Assert.Equal(expected, display);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_DURATION_DISPLAY_APPROXIMATED");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
        Assert.Equal(seconds / 86400, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Equal("0\"d\"", reopened.Sheets[0].CellAt(1, 1).GetStyle().NumberFormatCode);
        Assert.Equal(expected, Assert.Single(reopened.Sheets[0].Range("A1").CreateVisualSnapshot().Cells).Text);
        package.Position = 0;
        var read = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build().ReadDocument(package, "days.numbers");
        Assert.Equal(expected, Assert.Single(Assert.Single(read.Tables).Rows)[0]);
        Assert.Contains(read.Diagnostics, d => d.Code == "IWORK_READER_NUMBER_FORMAT_APPROXIMATED");
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Day_duration_is_editable_text_in_saved_document_tables(IWorkDocumentKind kind) {
        using var package = DurationFormatPackage(kind, -31017600d, DurationFormat(largest: 2, smallest: 2));
        var source = IWorkSourceDocument.Open(package, kind);
        using var saved = new MemoryStream();
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Word.WordDocument.Load(saved);
            Assert.Equal("-359d", reopened.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);
            Assert.Empty(reopened.ValidateDocument());
        } else {
            using var result = source.ToPowerPointPresentationResult(); Assert.False(result.IsVisualFallback);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.PowerPoint.PowerPointPresentation.Load(saved);
            Assert.Equal("-359d", reopened.Slides[0].Tables.Single().GetCell(0, 0).Text);
            Assert.Empty(reopened.ValidateDocument());
        }
    }

    [Fact]
    public void Native_day_durations_preserve_independent_caches_and_week_ranges_remain_unassessed() {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "numbers-parser");
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "duration-ranges.json")));
        foreach (var package in manifest.RootElement.GetProperty("packages").EnumerateArray()) {
            string path = Path.Combine(root, package.GetProperty("source").GetString()!);
            Assert.Equal(package.GetProperty("sourceSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant());
            var projection = IWorkSourceDocument.Open(path).ReadNumbers();
            foreach (var expected in package.GetProperty("cells").EnumerateArray()) {
                var table = projection.Sheets.Single(s => s.Name == expected.GetProperty("sheet").GetString())
                    .Tables.Single(t => t.Name == expected.GetProperty("table").GetString());
                var cell = table.GetCell(expected.GetProperty("row").GetInt32(), expected.GetProperty("column").GetInt32())!;
                Assert.Equal(expected.GetProperty("seconds").GetDouble(), cell.Value);
                if (expected.GetProperty("settings").GetProperty("duration_unit_largest").GetInt32() == 2) {
                    Assert.Equal(IWorkCellUnsupportedFeatures.None, cell.UnsupportedFeatures & IWorkCellUnsupportedFeatures.DurationFormat);
                    Assert.True(cell.TryGetFormattedNumber(out string display, out _));
                    Assert.Equal(expected.GetProperty("independentDisplayText").GetString(), display);
                } else {
                    Assert.Null(cell.NumberFormat);
                    Assert.Equal(IWorkCellUnsupportedFeatures.DurationFormat, cell.UnsupportedFeatures & IWorkCellUnsupportedFeatures.DurationFormat);
                }
            }
        }
    }
}
