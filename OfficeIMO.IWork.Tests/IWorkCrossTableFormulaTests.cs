using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(0, true)]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(3, false)]
    [InlineData(4, false)]
    [InlineData(5, false)]
    [InlineData(6, false)]
    public void Qualified_merge_table_identity_must_match_without_ambiguous_uuid_fields(int variant, bool expected) {
        const string tableId = "00112233-4455-6677-8899-aabbccddeeff";
        byte[] firstWord = variant == 1 ? VarintField(2, 0x33221101)
            : variant == 3 ? BytesField(2, new byte[] { 1 })
            : variant == 5 ? VarintField(2, (ulong)uint.MaxValue + 1)
            : VarintField(2, 0x33221100);
        byte[] uuid = Message(firstWord, VarintField(3, 0x77665544), VarintField(4, 0xbbaa9988),
            variant == 4 ? Array.Empty<byte>() : VarintField(5, 0xffeeddcc),
            variant == 2 ? VarintField(2, 0x33221100) : Array.Empty<byte>());
        byte[] range = Message(BytesField(3, Message(VarintField(1, 0), VarintField(2, 1))),
            BytesField(4, Message(VarintField(1, 0), VarintField(2, 1))));
        byte[] node = Message(VarintField(1, 67), BytesField(40, range), BytesField(28, BytesField(1, uuid)));
        var formula = IWorkProtobuf.Parse(BytesField(1, BytesField(1, node)), new IWorkReadOptions());
        var table = IWorkProtobuf.Parse(Message(StringField(1, tableId),
            variant == 6 ? StringField(1, tableId) : Array.Empty<byte>()), new IWorkReadOptions());
        Assert.Equal(expected, IWorkFormulaReader.TryReadAbsoluteRange(formula, 10,
            out _, out _, out _, out _, table));
        Assert.False(IWorkFormulaReader.Render(formula, 0, 0, 10, 100).IsComplete);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Cross_table_merge_declarations_retain_evidence_and_unmerged_source_cells(IWorkDocumentKind kind) {
        byte[] range = Message(BytesField(3, Message(VarintField(1, 0), VarintField(2, 1))),
            BytesField(4, Message(VarintField(1, 0), VarintField(2, 1))));
        byte[] node = Message(VarintField(1, 67), BytesField(40, range), BytesField(28, Array.Empty<byte>()));
        byte[] pair = BytesField(3, BytesField(2, BytesField(1, BytesField(1, node))));
        using MemoryStream package = MergeTestPackage(kind, MergeTestOwner(pair));
        var projection = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind);
        Assert.Empty(projection.Item1.MergedRanges);
        Assert.Equal(42d, Assert.Single(projection.Item1.Cells).Value);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind);
        IWorkSourceDeclarationIssue issue = Assert.Single(report.SourceDeclarationIssues);
        Assert.Equal("47/2/3[1]/2", issue.FieldPath);
        Assert.Equal(IWorkSourceDeclarationIssueKind.RejectedMessageSet, issue.Kind);
    }

    [Theory]
    [InlineData(false, 0)]
    [InlineData(false, 1)]
    [InlineData(false, 2)]
    [InlineData(true, 0)]
    [InlineData(true, 1)]
    [InlineData(true, 2)]
    public void Cross_table_ranges_are_incomplete_and_cannot_define_local_merges(bool tract, int representation) {
        byte[] external = representation switch {
            0 => BytesField(28, Message(BytesField(1, Message(VarintField(1, 123))))),
            1 => VarintField(28, 123),
            _ => Message(BytesField(28, Array.Empty<byte>()), BytesField(28, Array.Empty<byte>()))
        };
        byte[] nodes;
        if (tract) {
            byte[] range = Message(BytesField(3, Message(VarintField(1, 0), VarintField(2, 1))),
                BytesField(4, Message(VarintField(1, 0), VarintField(2, 1))));
            nodes = BytesField(1, Message(VarintField(1, 67), BytesField(40, range), external));
        } else {
            byte[] coordinate = Message(VarintField(1, 0), VarintField(2, 1));
            byte[] first = Message(VarintField(1, 36), BytesField(26, coordinate), BytesField(27, coordinate), external);
            byte[] lastCoordinate = Message(VarintField(1, 2), VarintField(2, 1));
            byte[] last = Message(VarintField(1, 36), BytesField(26, lastCoordinate), BytesField(27, lastCoordinate));
            nodes = Message(BytesField(1, first), BytesField(1, last), BytesField(1, VarintField(1, 29)));
        }
        var formula = IWorkProtobuf.Parse(BytesField(1, nodes), new IWorkReadOptions());
        Assert.False(IWorkFormulaReader.Render(formula, 0, 0, 10, 100).IsComplete);
        Assert.False(IWorkFormulaReader.TryReadAbsoluteRange(formula, 10, out _, out _, out _, out _));
    }
}

public sealed class IWorkCrossTableFormulaCorpusTests {
    private static string Fixture(string extension) => Path.Combine(AppContext.BaseDirectory,
        "Documents", "IWorkCorpus", "numbers-parser", "cross-table-formulas." + extension);

    [Fact]
    public void Independent_cross_table_rectangles_resolve_identity_and_preserve_mixed_endpoints() {
        using var manifest = JsonDocument.Parse(File.ReadAllText(Fixture("json")));
        string hash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Fixture("numbers")))).ToLowerInvariant();
        Assert.Equal(manifest.RootElement.GetProperty("sourceSha256").GetString(), hash);
        var source = IWorkSourceDocument.Open(Fixture("numbers"));
        IWorkNumbersProjection projection = source.ReadNumbers();
        foreach (JsonElement expected in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            IWorkTable table = Assert.Single(Assert.Single(projection.Sheets,
                s => s.Name == expected.GetProperty("sourceSheet").GetString()).Tables,
                t => t.Name == expected.GetProperty("sourceTable").GetString());
            IWorkTableCell cell = Assert.Single(table.Cells, c => c.Row == expected.GetProperty("row").GetInt32()
                && c.Column == expected.GetProperty("column").GetInt32());
            Assert.Equal(IWorkCellKind.Formula, cell.Kind);
            Assert.True(cell.FormulaIsComplete);
            Assert.Equal("=COUNTA('Main Sheet'::'Food Table'::" + expected.GetProperty("referenceAddress").GetString() + ")", cell.Formula);
            Assert.Equal(expected.GetProperty("cachedValue").GetDouble(), Assert.IsType<double>(cell.Value));
            Assert.True(cell.CachedValueIsComplete);
        }
    }

    [Fact]
    public void Saved_cross_table_formulas_use_actual_worksheet_names_and_preserve_caches() {
        using var manifest = JsonDocument.Parse(File.ReadAllText(Fixture("json")));
        using var converted = ExcelIWorkConverter.ConvertNumbersToExcelResult(Fixture("numbers"),
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true, NormalizeWorksheetNames = true });
        Assert.False(converted.IsVisualFallback);
        Assert.Throws<InvalidOperationException>(() => converted.Report.RequireNoLoss());
        using var saved = new MemoryStream(); converted.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        foreach (JsonElement expected in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            string sheetName = Assert.Single(converted.WorksheetMappings, m =>
                m.SourceSheetName == expected.GetProperty("sourceSheet").GetString()
                && m.SourceTableName == expected.GetProperty("sourceTable").GetString()).DestinationName;
            ExcelSheet sheet = Assert.Single(reopened.Sheets, s => s.Name == sheetName);
            ExcelCellData cell = sheet.CellAt(expected.GetProperty("row").GetInt32(),
                expected.GetProperty("column").GetInt32()).GetValue();
            IWorkTable table = Assert.Single(Assert.Single(converted.Projection.Sheets,
                s => s.Name == expected.GetProperty("sourceSheet").GetString()).Tables,
                t => t.Name == expected.GetProperty("sourceTable").GetString());
            IWorkFormulaCellStatus status = Assert.Single(converted.Report.FormulaCells, f =>
                f.TableIdentity?.RecordIdentifier == table.SourceIdentity!.RecordIdentifier
                && f.Row == expected.GetProperty("row").GetInt32() && f.Column == expected.GetProperty("column").GetInt32());
            Assert.True(status.ExpressionIsAssessed);
            Assert.True(status.ExpressionIsComplete);
            Assert.Equal(IWorkFormulaCacheStatus.Complete, status.CacheStatus);
            string target = Assert.Single(converted.WorksheetMappings, m =>
                m.SourceSheetName == expected.GetProperty("targetSheet").GetString()
                && m.SourceTableName == expected.GetProperty("targetTable").GetString()).DestinationName;
            Assert.Equal("COUNTA('" + target.Replace("'", "''") + "'!" + expected.GetProperty("referenceAddress").GetString() + ")",
                sheet.GetFormulaText(expected.GetProperty("row").GetInt32(), expected.GetProperty("column").GetInt32()));
            Assert.Equal(ExcelCellDataKind.Formula, cell.Kind);
            Assert.Equal(expected.GetProperty("cachedValue").GetDouble(), Assert.IsType<double>(cell.Value));
        }
    }
}
