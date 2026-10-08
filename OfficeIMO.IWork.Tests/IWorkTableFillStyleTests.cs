using System.Text.Json;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages, false)]
    [InlineData(IWorkDocumentKind.Numbers, false)]
    [InlineData(IWorkDocumentKind.Keynote, false)]
    [InlineData(IWorkDocumentKind.Pages, true)]
    [InlineData(IWorkDocumentKind.Numbers, true)]
    [InlineData(IWorkDocumentKind.Keynote, true)]
    public void Role_and_banded_fills_preserve_sparse_defaults_and_selected_overrides_in_saved_destinations(IWorkDocumentKind kind, bool banded) {
        using var package = RoleFillPackage(kind, banded: banded);
        var source = IWorkSourceDocument.Open(package);
        var policy = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        using var saved = new MemoryStream();
        IWorkTable table;
        string?[] expected = { "FF0000", "FFFF00", banded ? "808080" : "0000FF", "FFFF00", null, "0000FF" };
        (int Row, int Column)[] positions = { (1, 1), (4, 1), (3, 2), (4, 2), (2, 2), (2, 3) };
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(policy);
            Assert.False(result.IsVisualFallback); table = Assert.Single(result.Projection.Tables);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Word.WordDocument.Load(saved);
            for (int i = 0; i < positions.Length; i++) {
                var (row, column) = positions[i]; var cell = reopened.Tables[0].Rows[row - 1].Cells[column - 1];
                if (expected[i] == null) Assert.Equal(OfficeIMO.Word.WordShadingPattern.Nil, cell.ShadingPattern);
                else Assert.Equal(expected[i], cell.ShadingFillColorHex);
            }
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var result = source.ToExcelDocumentResult(policy);
            Assert.False(result.IsVisualFallback); table = Assert.Single(Assert.Single(result.Projection.Sheets).Tables);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
            for (int i = 0; i < positions.Length; i++) {
                var (row, column) = positions[i];
                Assert.Equal(expected[i] == null ? null : "FF" + expected[i], reopened.Sheets[0].GetCellStyle(row, column).FillColorArgb);
            }
            Assert.True(reopened.Sheets[0].TryGetCellValueSnapshot(3, 2, out var value));
            Assert.Equal(OfficeIMO.Excel.ExcelCellValueKind.Number, value!.Kind); Assert.Equal("42", value.RawValue);
        } else {
            using var result = source.ToPowerPointPresentationResult(policy);
            Assert.False(result.IsVisualFallback); table = Assert.Single(Assert.Single(result.Projection.Slides).Tables);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.PowerPoint.PowerPointPresentation.Load(saved);
            for (int i = 0; i < positions.Length; i++) {
                var (row, column) = positions[i]; var cell = Assert.Single(reopened.Slides[0].Tables).GetCell(row - 1, column - 1);
                Assert.Equal(expected[i], cell.FillColor); Assert.Equal(expected[i] == null, cell.NoFill);
            }
            Assert.Empty(reopened.ValidateDocument());
        }
        Assert.Equal(2, table.Cells.Count); // Region defaults do not allocate twelve source cells.
        Assert.Null(table.GetCell(1, 1)); Assert.Equal(42d, table.GetCell(3, 2)!.Value);
        for (int i = 0; i < positions.Length; i++) Assert.Equal(expected[i], table.GetFill(positions[i].Row, positions[i].Column)!.Color?.RgbHex);
        Assert.Throws<ArgumentOutOfRangeException>(() => table.GetFill(0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => table.GetFill(1, 4));
    }

    [Theory]
    [InlineData("missing-band-fill")]
    [InlineData("band-gradient")]
    [InlineData("duplicate-band-fill")]
    [InlineData("bad-banding")]
    [InlineData("missing-role")]
    [InlineData("wrong-role-type")]
    [InlineData("role-cycle")]
    [InlineData("role-gradient")]
    [InlineData("duplicate-role")]
    [InlineData("unresolved-selected")]
    public void Unsupported_role_or_selected_fills_are_reported_without_false_default_fallback(string defect) {
        using var package = RoleFillPackage(IWorkDocumentKind.Pages, defect);
        var projection = IWorkSourceDocument.Open(package).ReadPages(); var table = Assert.Single(projection.Tables);
        Assert.False(projection.HasEditableContent);
        Assert.Equal(42d, table.GetCell(3, 2)!.Value);
        Assert.Contains(projection.Diagnostics, diagnostic => diagnostic.Code == (defect == "unresolved-selected"
            ? "IWORK_TABLE_CELL_FILL_UNSUPPORTED" : "IWORK_TABLE_FILL_DEFAULTS_UNSUPPORTED"));
        if (defect == "unresolved-selected") Assert.Null(table.GetFill(2, 2));
        else {
            Assert.Null(table.GetFill(3, 2));
            Assert.True(table.GetFill(2, 2)!.IsNone); // Supported explicit selection survives rejected defaults.
        }
        if (defect != "missing-band-fill") Assert.NotEmpty(projection.SourceDeclarationIssues.Concat<object>(projection.SourceReferenceIssues));
        if (defect == "wrong-role-type") {
            var issue = Assert.Single(projection.SourceReferenceIssues);
            Assert.Equal(IWorkSourceReferenceIssueKind.UnexpectedTargetType, issue.Kind);
            Assert.Equal(11ul, issue.Owner.RecordIdentifier);
            Assert.Equal("18", issue.FieldPath);
            Assert.Equal(41ul, issue.TargetIdentifier);
        }
    }

    [Fact]
    public void Numbers_role_defaults_respect_the_bounded_destination_style_budget() {
        using var package = RoleFillPackage(IWorkDocumentKind.Numbers, rows: 100_001);
        using var result = IWorkSourceDocument.Open(package).ToExcelDocumentResult(
            IWorkTestPolicy.ForIncompletePreview(new IWorkConversionOptions { AllowPartialEditableReconstruction = true }));
        Assert.True(result.IsVisualFallback);
        var table = Assert.Single(Assert.Single(result.Projection.Sheets).Tables);
        Assert.Equal(2, table.Cells.Count);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Message.Contains("styled-cell budget", StringComparison.Ordinal));
    }

    [Fact]
    public void Native_unbanded_Keynote_defaults_suppress_destination_theme_fills_after_save() {
        using var manifest = JsonDocument.Parse(File.ReadAllText(CorpusFixture("keynote-table-fill-defaults.json")));
        var native = manifest.RootElement.GetProperty("source");
        string path = CorpusFixture(native.GetProperty("path").GetString()!);
        Assert.Equal(native.GetProperty("sha256").GetString(), Convert.ToHexString(
            System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant());
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(path,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.False(result.IsVisualFallback);
        var expected = Assert.Single(manifest.RootElement.GetProperty("tables").EnumerateArray());
        var table = Assert.Single(result.Projection.Slides.SelectMany(slide => slide.Tables));
        Assert.Equal(expected.GetProperty("modelIdentifier").GetUInt64(), table.ModelRecord!.Identifier);
        Assert.True(table.FillStyles.Body!.IsNone); Assert.True(table.FillStyles.HeaderRow!.IsNone);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.PowerPoint.PowerPointPresentation.Load(saved);
        var destination = Assert.Single(reopened.Slides.SelectMany(slide => slide.Tables));
        foreach (var cell in expected.GetProperty("cells").EnumerateArray()) {
            int row = cell.GetProperty("row").GetInt32(), column = cell.GetProperty("column").GetInt32();
            Assert.True(table.GetFill(row, column)!.IsNone); Assert.True(destination.GetCell(row - 1, column - 1).NoFill);
        }
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_FILL_DEFAULTS_UNSUPPORTED");
        Assert.Empty(reopened.ValidateDocument());
    }

    private static MemoryStream RoleFillPackage(IWorkDocumentKind kind, string? defect = null, int rows = 4, int columns = 3, bool banded = false, bool roleDefaults = true) {
        var records = new List<byte[]>();
        if (kind == IWorkDocumentKind.Pages) {
            records.Add(ArchiveRecord(1, 10000, ReferenceField(4, 2), new ulong[] { 2, 10 }));
            records.Add(ArchiveRecord(2, 2001, StringField(3, "Body")));
        } else if (kind == IWorkDocumentKind.Numbers) {
            records.Add(ArchiveRecord(1, 1, ReferenceField(1, 2)));
            records.Add(ArchiveRecord(2, 2, Message(StringField(1, "Sheet"), ReferenceField(2, 10))));
        } else {
            records.Add(ArchiveRecord(1, 1, ReferenceField(2, 2)));
            records.Add(ArchiveRecord(2, 2, KeynoteShow(ReferenceField(2, 3))));
            records.Add(ArchiveRecord(3, 4, ReferenceField(2, 4)));
            records.Add(ArchiveRecord(4, 5, ReferenceField(6, 10)));
        }
        records.Add(ArchiveRecord(10, 6000, Message(ReferenceField(2, 11),
            kind == IWorkDocumentKind.Keynote ? BytesField(1, GeometryDrawable(0, 0, 300, 120)) : Array.Empty<byte>())));
        var role = ReferenceField(18, defect == "missing-role" ? 99UL : 41UL);
        byte[] store = Message(ReferenceField(5, 13), BytesField(3, BytesField(1, Message(VarintField(1, 0), ReferenceField(2, 12)))));
        records.Add(ArchiveRecord(11, 6001, Message(VarintField(6, (ulong)rows), VarintField(7, (ulong)columns),
            VarintField(9, 1), VarintField(10, 1), VarintField(11, 1), BytesField(4, store), ReferenceField(3, 40), roleDefaults ? role : Array.Empty<byte>(),
            defect == "duplicate-role" ? role : Array.Empty<byte>(),
            roleDefaults ? Message(ReferenceField(19, 42), ReferenceField(20, 43), ReferenceField(21, 44)) : Array.Empty<byte>())));
        byte[] empty = new byte[16]; empty[0] = 5; empty[8] = 0x20; empty[12] = 1;
        byte[] inherited = (byte[])empty.Clone(); inherited[12] = 2;
        byte[] number = new byte[20]; number[0] = 5; number[1] = 2; number[8] = 2; Buffer.BlockCopy(BitConverter.GetBytes(42d), 0, number, 12, 8);
        records.Add(ArchiveRecord(12, 6002, Message(VarintField(1, 2), VarintField(2, 2), VarintField(3, 3), VarintField(4, 2),
            BytesField(5, Message(VarintField(1, 1), VarintField(2, 2), BytesField(6, Message(empty, inherited)), BytesField(7, new byte[] { 255, 255, 0, 0, 16, 0 }))),
            BytesField(5, Message(VarintField(1, 2), VarintField(2, 1), BytesField(6, number), BytesField(7, new byte[] { 255, 255, 0, 0 }))),
            VarintField(6, 5), VarintField(7, 1))));
        records.Add(ArchiveRecord(13, defect == "wrong-catalog-target-type" ? 2021u : 6005u, Message(VarintField(1, 4),
            BytesField(3, Message(VarintField(1, 1), ReferenceField(4, defect == "unresolved-selected" ? 99UL : 30UL))),
            BytesField(3, Message(VarintField(1, 2), ReferenceField(4, 31))))));
        records.Add(FillStyle(30, Array.Empty<byte>())); records.Add(FillStyle(31, null));
        bool activeBand = banded || defect is "missing-band-fill" or "band-gradient" or "duplicate-band-fill" or "inherited-band";
        byte[] bandFill = BytesField(2, defect == "band-gradient" ? BytesField(2, Message()) : FillColor(0.5f, 0.5f, 0.5f));
        records.Add(ArchiveRecord(40, 6003, Message(BytesField(1, ReferenceField(3, 45)),
            BytesField(11, Message(VarintField(1, defect == "bad-banding" ? 2UL : activeBand ? 1UL : 0UL),
                activeBand && defect is not ("missing-band-fill" or "inherited-band") ? bandFill : Array.Empty<byte>(),
                defect == "duplicate-band-fill" ? bandFill : defect == "inactive-band-gradient" ? BytesField(2, BytesField(2, Message())) : Array.Empty<byte>())))));
        records.Add(ArchiveRecord(45, 6003, BytesField(11, Message(VarintField(1, 1),
            defect == "inherited-band" ? bandFill : Array.Empty<byte>())))); // Child can disable inherited banding.
        records.Add(defect == "wrong-role-type" ? ArchiveRecord(41, 2022, Message())
            : FillStyle(41, defect == "role-gradient" ? BytesField(2, Message()) : null, defect == "role-cycle" ? 41UL : 46UL));
        records.Add(FillStyle(46, FillColor(0, 0, 1)));
        records.Add(FillStyle(42, FillColor(1, 0, 0))); records.Add(FillStyle(43, FillColor(0, 1, 0))); records.Add(FillStyle(44, FillColor(1, 1, 0)));
        return CreatePackage(("Index/Document.iwa", FrameIwa(Message(records.ToArray()))), ("preview.png", ValidPreviewPng()));
    }
}
