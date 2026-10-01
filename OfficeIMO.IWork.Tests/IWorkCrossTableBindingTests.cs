using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Forward_cross_table_binding_uses_normalized_unique_names_and_preserves_numeric_cache() {
        using MemoryStream package = CrossBindingPackage();
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions { NormalizeWorksheetNames = true });
        Assert.False(result.IsVisualFallback);
        Assert.DoesNotContain(result.Projection.Diagnostics, d => d.Code == "IWORK_TABLE_FORMULA_PARTIAL");
        IWorkTableCell source = result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!;
        Assert.True(source.FormulaIsComplete);
        Assert.Equal(13d, source.Value);
        string target = result.WorksheetMappings[1].DestinationName;
        Assert.NotEqual(target, result.WorksheetMappings[2].DestinationName);
        Assert.True(result.WorksheetMappings[1].WasRenamed);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal("SUM('" + target.Replace("'", "''") + "'!$A$1:$B$2)", reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(13d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Equal(13d, reopened.Sheets[1].CellAt(1, 1).GetValue<double>());
        Assert.Same(result.Projection.Sheets[0].Tables[0], result.Projection.Sheets[0].Drawables[0].Table);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Ambiguous_missing_and_inactive_target_identities_never_bind_a_cached_formula(int defect) {
        using MemoryStream package = CrossBindingPackage(defect);
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions { NormalizeWorksheetNames = true });
        IWorkTableCell source = result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!;
        Assert.False(source.FormulaIsComplete);
        Assert.Equal(13d, source.Value);
        Assert.Contains(result.Projection.Diagnostics, d => d.Code == "IWORK_TABLE_FORMULA_PARTIAL");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Null(reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(13d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
    }

    [Fact]
    public void Bound_cross_table_expression_without_cache_remains_editable() {
        using MemoryStream package = CrossBindingPackage(uncached: true);
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers).ReadNumbers();
        Assert.True(projection.HasEditableContent);
        IWorkTableCell cell = projection.Sheets[0].Tables[0].GetCell(1, 1)!;
        Assert.True(cell.FormulaIsComplete);
        Assert.Null(cell.Value);
        Assert.DoesNotContain(projection.Diagnostics, d => d.Code == "IWORK_TABLE_FORMULA_UNSUPPORTED");
        package.Position = 0;
        using var converted = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions { NormalizeWorksheetNames = true });
        Assert.False(converted.IsVisualFallback);
        Assert.NotNull(converted.Value.Sheets[0].GetFormulaText(1, 1));
    }

    [Theory]
    [InlineData(4, false)]
    [InlineData(8, true)]
    public void Source_and_destination_binding_charge_cumulative_formula_work(long maximumOperations, bool destination) {
        using MemoryStream package = CrossBindingPackage();
        var options = new IWorkReadOptions { MaximumFormulaRenderingOperations = maximumOperations };
        if (destination) Assert.Throws<InvalidDataException>(() => ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            options, new IWorkConversionOptions { NormalizeWorksheetNames = true }));
        else Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers, options).ReadNumbers());
    }

    [Fact]
    public void Qualified_source_formula_text_consumes_projection_character_budget() {
        using MemoryStream package = CrossBindingPackage(shortNames: true);
        Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumProjectedTextCharacters = 64 }).ReadNumbers());
        package.Position = 0;
        Assert.True(IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumProjectedTextCharacters = 4096 }).ReadNumbers().HasEditableContent);
    }

    [Fact]
    public void Destination_name_expansion_respects_formula_limits_and_discards_planned_worksheets_on_fallback() {
        using MemoryStream package = CrossBindingPackage(shortNames: true, unnamedTarget: true);
        var read = new IWorkReadOptions { MaximumFormulaCharacters = 30 };
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package, read);
        Assert.True(result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!.FormulaIsComplete);
        Assert.True(result.IsVisualFallback);
        Assert.Empty(result.WorksheetMappings);
        Assert.Equal("Preview", Assert.Single(result.Value.Sheets).Name);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_EXCEL_DESTINATION_UNSUPPORTED");
        package.Position = 0;
        Assert.Throws<NotSupportedException>(() => ExcelIWorkConverter.ConvertNumbersToExcelResult(package, read,
            new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly }));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Rectangular_reference_rejects_ambiguous_sticky_flags_and_invalid_relative_coordinates(int defect) {
        byte[] sticky = defect == 0
            ? Message(VarintField(1, 0), VarintField(1, 1))
            : defect == 1 ? BytesField(1, new byte[] { 1 }) : Array.Empty<byte>();
        ulong offset = defect == 2 ? 0x100000000UL : 0;
        byte[] range = Message(BytesField(1, VarintField(1, offset)), BytesField(2, VarintField(1, 0)));
        byte[] node = Message(VarintField(1, 67), BytesField(33, sticky), BytesField(40, range));
        var formula = IWorkProtobuf.Parse(BytesField(1, BytesField(1, node)), new IWorkReadOptions());
        Assert.False(IWorkFormulaReader.Render(formula, 0, 0, 10, 100).IsComplete);
    }

    private static MemoryStream CrossBindingPackage(int defect = 0, bool uncached = false, bool shortNames = false, bool unnamedTarget = false) {
        const string targetId = "00112233-4455-6677-8899-aabbccddeeff";
        byte[] uuid = Message(VarintField(2, 0x33221100), VarintField(3, 0x77665544),
            VarintField(4, 0xbbaa9988), VarintField(5, 0xffeeddcc));
        byte[] range = Message(BytesField(3, Message(VarintField(1, 0), VarintField(2, 1))),
            BytesField(4, Message(VarintField(1, 0), VarintField(2, 1))));
        byte[] reference = Message(VarintField(1, 67), BytesField(40, range), BytesField(28, BytesField(1, uuid)));
        byte[] formula = BytesField(1, Message(BytesField(1, reference),
            BytesField(1, Message(VarintField(1, 16), VarintField(2, 168), VarintField(3, 1)))));
        var records = new List<byte[]> { ArchiveRecord(1, 1, ReferenceField(1, 2)) };
        var sheet = new List<byte[]> { StringField(1, shortNames ? "Sheet" : "Quarter's /") };
        for (int index = 0; index < 3; index++) {
            ulong info = (ulong)(10 + index * 4), model = info + 1, tile = info + 2;
            if (index != 1 || defect != 3) sheet.Add(ReferenceField(2, info));
            string id = index == 1 ? (defect == 2 ? "unknown" : targetId)
                : index == 2 && defect == 1 ? targetId : $"00000000-0000-0000-0000-{index + 1:000000000000}";
            string name = shortNames ? index == 0 ? "Source" : "Target" + index
                : index == 0 ? "Source" : index == 1 ? "Totals /" : "Totals ?";
            if (index == 1 && unnamedTarget) name = string.Empty;
            var spec = new TableSpec(name, 2, 2, 13d, hasFormula: index == 0,
                formulaWithoutCachedValue: uncached && index == 0);
            byte[] store = BytesField(3, BytesField(1, Message(VarintField(1, 0), ReferenceField(2, tile))));
            if (index == 0) store = Message(store, ReferenceField(6, 30));
            records.Add(ArchiveRecord(info, 6000, ReferenceField(2, model)));
            records.Add(ArchiveRecord(model, 6001, Message(StringField(1, id), StringField(8, name),
                VarintField(6, 2), VarintField(7, 2), BytesField(4, store))));
            records.Add(ArchiveRecord(tile, 6002, BytesField(5, CreateBncRow(spec))));
        }
        records.Add(ArchiveRecord(30, 6201, BytesField(3, Message(VarintField(1, 0), BytesField(5, formula)))));
        records.Add(ArchiveRecord(2, 2, Message(sheet.ToArray())));
        return CreatePackage(("Index/Document.iwa", FrameIwa(Message(records.ToArray()))), ("preview.png", ValidPreviewPng()));
    }
}
