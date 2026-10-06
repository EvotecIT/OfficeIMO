using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using OfficeIMO.IWork;
using OfficeIMO.Reader.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Independent_native_user_hidden_columns_survive_uuid_permutations_and_saved_xlsx() {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus");
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "numbers-parser", "user-hidden-columns.json")));
        string path = Path.Combine(root, manifest.RootElement.GetProperty("source").GetProperty("path").GetString()!);
        Assert.Equal(manifest.RootElement.GetProperty("source").GetProperty("sha256").GetString(),
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant());
        IWorkSourceDocument source = IWorkSourceDocument.Open(path);
        using var result = source.ToExcelDocumentResult(new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        IWorkTable[] tables = result.Projection.Sheets.SelectMany(sheet => sheet.Tables).ToArray();
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
        saved.Position = 0;
        using var archive = new ZipArchive(saved, ZipArchiveMode.Read, leaveOpen: true);
        foreach (JsonElement expected in manifest.RootElement.GetProperty("tables").EnumerateArray()) {
            Assert.Equal(6267u, Assert.Single(source.Records, record => record.Identifier == expected.GetProperty("mapIdentifier").GetUInt64()
                && record.PayloadIndex == 0).MessageType);
            int ordinal = Array.FindIndex(tables, table => table.ModelRecord!.Identifier == expected.GetProperty("modelIdentifier").GetUInt64());
            Assert.True(ordinal >= 0);
            IWorkTable table = tables[ordinal];
            int[] columns = expected.GetProperty("hiddenColumns").EnumerateArray().Select(column => column.GetInt32()).ToArray();
            Assert.Equal(columns, table.HiddenColumns);
            Assert.Empty(table.HiddenRows);
            Assert.Throws<NotSupportedException>(() => ((IList<int>)table.HiddenColumns).Clear());
            using Stream xml = archive.GetEntry($"xl/worksheets/sheet{ordinal + 1}.xml")!.Open();
            XElement worksheet = XElement.Load(xml);
            for (int column = 1; column <= table.ColumnCount; column++) {
                bool hidden = worksheet.Descendants(SpreadsheetNs + "col").Any(element =>
                    (int)element.Attribute("min")! <= column && (int)element.Attribute("max")! >= column
                    && (string?)element.Attribute("hidden") is "1" or "true");
                Assert.Equal(columns.Contains(column), hidden);
            }
        }
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_HIDDEN_STATES_UNASSESSED");
    }

    [Fact]
    public void Qualified_hidden_rows_and_columns_keep_populated_cells_and_reader_content() {
        using MemoryStream package = MappedHiddenPackage(IWorkDocumentKind.Numbers);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package);
        using var result = source.ToExcelDocumentResult();
        Assert.False(result.IsVisualFallback);
        Assert.False(result.Report.IsPartialEditableReconstruction);
        IWorkTable table = result.Projection.Sheets[0].Tables[0];
        Assert.Equal(new[] { 1 }, table.HiddenRows);
        Assert.Equal(new[] { 1 }, table.HiddenColumns);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
        Assert.Equal(42d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        saved.Position = 0;
        using var archive = new ZipArchive(saved, ZipArchiveMode.Read, leaveOpen: true);
        using Stream xml = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        XElement worksheet = XElement.Load(xml);
        Assert.Equal("1", (string?)worksheet.Descendants(SpreadsheetNs + "row").Single(row => (int)row.Attribute("r")! == 1).Attribute("hidden"));
        Assert.Equal("1", (string?)Assert.Single(worksheet.Descendants(SpreadsheetNs + "col")).Attribute("hidden"));
        package.Position = 0;
        var read = new OfficeIMO.Reader.OfficeDocumentReaderBuilder().AddIWorkHandler().Build().ReadDocument(package, "hidden.numbers");
        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "IWORK_READER_HIDDEN_TABLE_CONTENT_INCLUDED");
        Assert.Equal("42", Assert.Single(Assert.Single(read.Tables).Rows[0]));
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Text_table_destinations_require_partial_policy_for_qualified_visibility(IWorkDocumentKind kind) {
        using MemoryStream package = MappedHiddenPackage(kind);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind);
        if (kind == IWorkDocumentKind.Pages) {
            using var automatic = source.ToWordDocumentResult(options:new IWorkConversionOptions { RequireCompleteVisualCoverage = false }); Assert.True(automatic.IsVisualFallback);
        } else {
            using var automatic = source.ToPowerPointPresentationResult(options:new IWorkConversionOptions { RequireCompleteVisualCoverage = false }); Assert.True(automatic.IsVisualFallback);
        }
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual: false);
        Assert.True(report.IsPartialEditableReconstruction);
        Assert.Throws<InvalidOperationException>(() => report.RequireCompleteEditableReconstruction());
        Assert.Contains(report.FidelityDiagnostics, diagnostic => diagnostic.Code.EndsWith("TABLE_VISIBILITY_OMITTED", StringComparison.Ordinal)
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }

    [Theory]
    [InlineData("duplicateIndex")]
    [InlineData("wrongWire")]
    [InlineData("missingIndex")]
    [InlineData("inverseConflict")]
    [InlineData("duplicateUuid")]
    public void Untrusted_uuid_maps_retire_only_the_affected_axis(string defect) {
        byte[] rowUuids = Message(BytesField(4, VisibilityUuid(1)), BytesField(4, VisibilityUuid(defect == "duplicateUuid" ? 1ul : 2ul)),
            BytesField(4, VisibilityUuid(3)));
        byte[] indexes = defect switch {
            "duplicateIndex" => Message(VarintField(5, 2), VarintField(5, 0), VarintField(5, 0)),
            "wrongWire" => FloatField(5, 0),
            "missingIndex" => Message(VarintField(5, 2), VarintField(5, 0)),
            _ => Message(VarintField(5, 2), VarintField(5, 0), VarintField(5, 1))
        };
        byte[] inverse = defect == "inverseConflict" ? Message(VarintField(6, 0), VarintField(6, 1), VarintField(6, 2)) : Message();
        using MemoryStream package = MappedHiddenPackage(IWorkDocumentKind.Numbers, rowMap: Message(rowUuids, indexes, inverse));
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package).ReadNumbers();
        IWorkTable table = Assert.Single(Assert.Single(projection.Sheets).Tables);
        Assert.Empty(table.HiddenRows);
        Assert.Equal(new[] { 1 }, table.HiddenColumns);
        Assert.False(projection.HasEditableContent);
        Assert.Equal(42d, Assert.Single(table.Cells).Value);
        Assert.Contains(projection.SourceDeclarationIssues, issue => issue.Owner.RecordIdentifier == 30);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Duplicate_or_unresolved_base_state_identities_cannot_recover_hidden_positions(bool malformed) {
        byte[] row = BytesField(2, Message(BytesField(1, VisibilityUuid(2)), VarintField(2, 1)));
        byte[] conflict = BytesField(2, malformed ? new byte[] { 0x80 }
            : Message(BytesField(1, VisibilityUuid(2)), VarintField(2, 0)));
        using MemoryStream package = MappedHiddenPackage(IWorkDocumentKind.Numbers, rowStates: Message(row, conflict));
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package).ReadNumbers();
        Assert.Empty(Assert.Single(Assert.Single(projection.Sheets).Tables).HiddenRows);
        Assert.False(projection.HasEditableContent);
    }

    [Fact]
    public void Packed_mapping_indexes_are_supported_and_mapping_work_uses_existing_limits() {
        byte[] uuids = Message(BytesField(4, VisibilityUuid(1)), BytesField(4, VisibilityUuid(2)), BytesField(4, VisibilityUuid(3)));
        using MemoryStream package = MappedHiddenPackage(IWorkDocumentKind.Numbers, rowMap: Message(uuids, BytesField(5, new byte[] { 2, 0, 1 })));
        Assert.Equal(new[] { 1 }, Assert.Single(Assert.Single(IWorkSourceDocument.Open(package).ReadNumbers().Sheets).Tables).HiddenRows);
        package.Position = 0;
        Assert.Contains("limit", Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(package,
            new IWorkReadOptions { MaximumTableDimensionEntries = 10 }).ReadNumbers()).Message, StringComparison.OrdinalIgnoreCase);
    }

    private static byte[] VisibilityUuid(ulong lower, ulong upper = 10) => Message(VarintField(1, lower), VarintField(2, upper));

    private static MemoryStream MappedHiddenPackage(IWorkDocumentKind kind, byte[]? rowMap = null, byte[]? rowStates = null, uint mapType = 6267, bool repeatModel = false) =>
        HiddenStatePackage(kind, HiddenOwner(
            rowFields: rowStates ?? BytesField(2, Message(BytesField(1, VisibilityUuid(2)), VarintField(2, 1))),
            columnFields: BytesField(2, Message(BytesField(1, VisibilityUuid(100)), VarintField(2, 1)))),
            repeatModel: repeatModel, modelFields: ReferenceField(46, 30), records: ArchiveRecord(30, mapType, Message(
                BytesField(1, VisibilityUuid(100)), VarintField(2, 0), VarintField(3, 0),
                rowMap ?? Message(BytesField(4, VisibilityUuid(1)), BytesField(4, VisibilityUuid(2)), BytesField(4, VisibilityUuid(3)),
                    VarintField(5, 2), VarintField(5, 0), VarintField(5, 1), VarintField(6, 1), VarintField(6, 2), VarintField(6, 0)))));
}
