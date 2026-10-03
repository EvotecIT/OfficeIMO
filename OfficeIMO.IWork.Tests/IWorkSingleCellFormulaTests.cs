using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Single_cell_cross_table_binding_preserves_flags_forward_names_and_numeric_caches(bool rowAbsolute, bool columnAbsolute) {
        byte[] formula = SingleCellFormula(SingleCellReference(rowAbsolute, columnAbsolute));
        using MemoryStream package = CrossBindingPackage(formulaPayload: formula);
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions { NormalizeWorksheetNames = true });
        Assert.False(result.IsVisualFallback);
        IWorkTableCell cell = result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!;
        Assert.True(cell.FormulaIsComplete);
        string address = (columnAbsolute ? "$" : "") + "A" + (rowAbsolute ? "$" : "") + "1";
        string target = result.WorksheetMappings[1].DestinationName.Replace("'", "''");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal("'" + target + "'!" + address, reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(13d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_NUMBERS_TABLE_BODY_RANGE_APPROXIMATED");
        reopened.Sheets[1].CellValue(1, 1, 29d); reopened.Calculate();
        Assert.Equal(29d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Single_cell_missing_ambiguous_or_inactive_identities_retain_cache_without_local_formula(int defect) {
        using MemoryStream package = CrossBindingPackage(defect, formulaPayload: SingleCellFormula(SingleCellReference()));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions { NormalizeWorksheetNames = true });
        Assert.False(result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!.FormulaIsComplete);
        Assert.Null(result.Value.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(13d, result.Value.Sheets[0].CellAt(1, 1).GetValue<double>());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public void Single_cell_untrusted_coordinates_do_not_publish_editable_references(int defect) {
        byte[] coordinate = defect switch {
            0 => Message(VarintField(1, 0), VarintField(1, 2)),
            1 => BytesField(1, new byte[] { 0 }),
            2 => Message(VarintField(1, 0), VarintField(2, 2)),
            3 => VarintField(1, 1), // Relative -1 from source row zero.
            _ => VarintField(1, (ulong)int.MaxValue * 2 + 2)
        };
        using MemoryStream package = CrossBindingPackage(formulaPayload: SingleCellFormula(SingleCellReference(rowCoordinate: coordinate)));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions { NormalizeWorksheetNames = true });
        Assert.False(result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!.FormulaIsComplete);
        Assert.Null(result.Value.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(13d, result.Value.Sheets[0].CellAt(1, 1).GetValue<double>());
    }

    [Theory]
    [InlineData(0, true)]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(3, true)]
    public void Cross_table_endpoint_ranges_require_one_target_and_emit_one_qualifier(int variant, bool complete) {
        byte[] second = variant is 0 or 3 ? SingleCellIdentity() : variant == 1 ? Array.Empty<byte>()
            : BytesField(28, BytesField(1, Message(VarintField(2, 0), VarintField(3, 0), VarintField(4, 0), VarintField(5, 0x03000000))));
        byte[] last = Message(VarintField(1, 36), BytesField(26, Message(VarintField(1, 2), VarintField(2, 1))),
            BytesField(27, Message(VarintField(1, 2), VarintField(2, 1))), second);
        var nodes = new List<byte[]> { SingleCellReference(true, true) };
        if (variant == 3) nodes.Add(Message(VarintField(1, 32), StringField(25, " ")));
        nodes.Add(last);
        if (variant == 3) nodes.Add(Message(VarintField(1, 33), StringField(25, " ")));
        nodes.Add(VarintField(1, 29));
        nodes.Add(Message(VarintField(1, 16), VarintField(2, 168), VarintField(3, 1)));
        byte[] formula = SingleCellFormula(nodes.ToArray());
        using MemoryStream package = CrossBindingPackage(formulaPayload: formula);
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions { NormalizeWorksheetNames = true });
        Assert.Equal(complete, result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!.FormulaIsComplete);
        if (complete) {
            string target = result.WorksheetMappings[1].DestinationName.Replace("'", "''");
            Assert.Equal("SUM('" + target + "'!" + (variant == 3 ? "$A$1 : $B$2" : "$A$1:$B$2") + ")", result.Value.Sheets[0].GetFormulaText(1, 1));
        } else Assert.Null(result.Value.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(13d, result.Value.Sheets[0].CellAt(1, 1).GetValue<double>());
    }

    [Theory]
    [InlineData(27)]
    [InlineData(28)]
    [InlineData(63)]
    [InlineData(64)]
    [InlineData(65)]
    public void Other_cross_table_reference_node_families_remain_unqualified(int nodeType) {
        byte[] node = Message(VarintField(1, (ulong)nodeType), BytesField(26, VarintField(1, 0)),
            BytesField(27, VarintField(1, 0)), SingleCellIdentity());
        using MemoryStream package = CrossBindingPackage(formulaPayload: SingleCellFormula(node));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions { NormalizeWorksheetNames = true });
        Assert.False(result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!.FormulaIsComplete);
        Assert.Null(result.Value.Sheets[0].GetFormulaText(1, 1));
    }

    private static byte[] SingleCellFormula(params byte[][] nodes) => BytesField(1, Message(nodes.Select(node => BytesField(1, node)).ToArray()));
    private static byte[] SingleCellIdentity() => BytesField(28, BytesField(1, Message(VarintField(2, 0x33221100),
        VarintField(3, 0x77665544), VarintField(4, 0xbbaa9988), VarintField(5, 0xffeeddcc))));
    private static byte[] SingleCellReference(bool rowAbsolute = false, bool columnAbsolute = false, byte[]? rowCoordinate = null) =>
        Message(VarintField(1, 36), BytesField(26, Message(VarintField(1, 0), VarintField(2, columnAbsolute ? 1ul : 0ul))),
            BytesField(27, rowCoordinate ?? Message(VarintField(1, 0), VarintField(2, rowAbsolute ? 1ul : 0ul))), SingleCellIdentity());
}
