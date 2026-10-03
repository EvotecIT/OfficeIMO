using System.Xml.Linq;
using System.Text.Json;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Numbers_formula_binding_preserves_native_table_row_sizing() {
        var source = IWorkSourceDocument.Open(CorpusFixture("numbers-parser/test-10-formulas.numbers"));
        IWorkTable table = Assert.Single(source.ReadNumbers().Sheets.SelectMany(sheet => sheet.Tables),
            table => table.Name == "Table 2");
        Assert.Contains(table.Cells, cell => cell.Formula != null);
        Assert.True(table.AutoResizeRows);
    }

    [Fact]
    public void Independent_Pages_sizing_declarations_become_minimum_DOCX_heights() {
        using var manifest = JsonDocument.Parse(File.ReadAllText(CorpusFixture("pages-table-sizing.json")));
        using var result = WordIWorkConverter.ConvertPagesToWordResult(CorpusFixture("picodocs/sample-v14.4.pages"),
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.False(result.IsVisualFallback);
        var expected = manifest.RootElement.GetProperty("tables").EnumerateArray().ToArray();
        Assert.Equal(expected.Length, result.Projection.Tables.Count);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Word.WordDocument.Load(saved);
        foreach (JsonElement native in expected) {
            string name = native.GetProperty("name").GetString()!;
            IWorkTable table = Assert.Single(result.Projection.Tables, table => table.Name == name);
            Assert.Equal(native.GetProperty("modelIdentifier").GetUInt64(), table.ModelRecord!.Identifier);
            Assert.Equal(native.GetProperty("autoResizeRows").GetBoolean(), table.AutoResizeRows);
            int position = result.Projection.Tables.ToList().IndexOf(table);
            foreach (JsonElement row in native.GetProperty("rowHeights").EnumerateArray()) {
                int number = row.GetProperty("row").GetInt32();
                double height = row.GetProperty("heightPoints").GetDouble();
                Assert.Equal(height, table.GetRowHeight(number));
                Assert.Equal((int)Math.Round(height * 20), reopened.Tables[position].Rows[number - 1].MinimumHeight);
            }
        }
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_ROW_SIZING_UNSUPPORTED");
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, true)]
    [InlineData(IWorkDocumentKind.Pages, false)]
    [InlineData(IWorkDocumentKind.Numbers, true)]
    [InlineData(IWorkDocumentKind.Keynote, true)]
    public void Native_table_auto_resize_survives_projection_and_Pages_row_constraints(
        IWorkDocumentKind kind, bool autoResize) {
        using MemoryStream package = DimensionPackage(kind, tableStyleReference: ReferenceField(3, 30),
            styleRecords: new[] { TableSizingStyle(30, autoResize) });
        IWorkSourceDocument source = IWorkSourceDocument.Open(package);
        IWorkTable table = kind switch {
            IWorkDocumentKind.Pages => source.ReadPages().Tables[0],
            IWorkDocumentKind.Numbers => source.ReadNumbers().Sheets[0].Tables[0],
            _ => source.ReadKeynote().Slides[0].Tables[0]
        };
        Assert.Equal(autoResize, table.AutoResizeRows);
        if (kind != IWorkDocumentKind.Pages) return;
        using var result = source.ToWordDocumentResult();
        Assert.False(result.IsVisualFallback);
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using var reopened = OfficeIMO.Word.WordDocument.Load(saved);
        Assert.Equal(new[] { 405, 200, 600 }, reopened.Tables[0].RowHeight);
        Assert.Equal(autoResize ? 405 : (int?)null, reopened.Tables[0].Rows[0].MinimumHeight);
        saved.Position = 0;
        using var zip = new ZipArchive(saved, ZipArchiveMode.Read, leaveOpen: true);
        using Stream xml = zip.GetEntry("word/document.xml")!.Open();
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        Assert.All(XDocument.Load(xml).Descendants(w + "trHeight"),
            height => Assert.Equal(autoResize ? "atLeast" : "exact", height.Attribute(w + "hRule")!.Value));
    }

    [Theory]
    [InlineData(null, true)]
    [InlineData(false, false)]
    public void Table_auto_resize_inherits_and_child_override_wins(bool? child, bool expected) {
        using MemoryStream package = DimensionPackage(IWorkDocumentKind.Pages,
            tableStyleReference: ReferenceField(3, 30),
            styleRecords: new[] { TableSizingStyle(30, child, 31), TableSizingStyle(31, true) });
        Assert.Equal(expected, IWorkSourceDocument.Open(package).ReadPages().Tables[0].AutoResizeRows);
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("wrong-type")]
    [InlineData("invalid-bool")]
    [InlineData("duplicate-bool")]
    [InlineData("duplicate-properties")]
    [InlineData("cycle")]
    [InlineData("depth")]
    [InlineData("missing-parent")]
    [InlineData("wrong-wire-bool")]
    [InlineData("malformed-properties")]
    [InlineData("duplicate-reference")]
    public void Unresolved_row_sizing_retains_evidence_and_prevents_complete_editable_output(string defect) {
        byte[] style = defect switch {
            "wrong-type" => ArchiveRecord(30, 2001, Message(StringField(3, "Not a style"))),
            "invalid-bool" => ArchiveRecord(30, 6003, Message(BytesField(11, Message(VarintField(22, 2))))),
            "duplicate-bool" => ArchiveRecord(30, 6003, Message(BytesField(11,
                Message(VarintField(22, 1), VarintField(22, 0))))),
            "duplicate-properties" => ArchiveRecord(30, 6003, Message(BytesField(11, Message(VarintField(22, 1))),
                BytesField(11, Message(VarintField(22, 0))))),
            "wrong-wire-bool" => ArchiveRecord(30, 6003, Message(BytesField(11, StringField(22, "true")))),
            "malformed-properties" => ArchiveRecord(30, 6003, Message(BytesField(11, new byte[] { 0xff }))),
            "cycle" => TableSizingStyle(30, true, 30),
            "missing-parent" => TableSizingStyle(30, true, 99),
            _ => TableSizingStyle(30, true, 31)
        };
        using MemoryStream package = DimensionPackage(IWorkDocumentKind.Pages,
            tableStyleReference: defect == "duplicate-reference"
                ? Message(ReferenceField(3, 30), ReferenceField(3, 31)) : ReferenceField(3, 30),
            styleRecords: defect == "missing" ? Array.Empty<byte[]>() : new[] { style, TableSizingStyle(31, false) });
        IWorkSourceDocument source = IWorkSourceDocument.Open(package,
            new IWorkReadOptions { MaximumTextStyleInheritanceDepth = defect == "depth" ? 1 : 128 });
        using var result = source.ToWordDocumentResult();
        Assert.True(result.IsVisualFallback);
        Assert.Null(result.Projection.Tables[0].AutoResizeRows);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_ROW_SIZING_UNSUPPORTED");
        Assert.Contains(result.Report.SourceDeclarationIssues, issue => issue.Owner.RecordIdentifier == 11 && issue.FieldPath == "3");
        if (defect == "missing") Assert.Contains(result.Report.SourceReferenceIssues,
            issue => issue.Owner.RecordIdentifier == 11 && issue.FieldPath == "3");
        if (defect is "invalid-bool" or "duplicate-bool" or "wrong-wire-bool" or "duplicate-properties" or "malformed-properties")
            Assert.Contains(result.Report.SourceDeclarationIssues, issue => issue.Owner.RecordIdentifier == 30
                && issue.FieldPath == (defect is "duplicate-properties" or "malformed-properties" ? "11" : "11/22"));
    }

    private static byte[] TableSizingStyle(ulong identifier, bool? autoResize, ulong? parent = null) =>
        ArchiveRecord(identifier, 6003, Message(
            parent.HasValue ? BytesField(1, ReferenceField(3, parent.Value)) : Array.Empty<byte>(),
            autoResize.HasValue ? BytesField(11, VarintField(22, autoResize.Value ? 1UL : 0UL)) : Array.Empty<byte>()));
}
