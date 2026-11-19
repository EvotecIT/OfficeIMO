using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Write;

namespace OfficeIMO.Excel.Xlsb.Write;

/// <summary>Encodes cell formulas of a newly generated workbook against its defined names and self-reference table.</summary>
/// <remarks>
/// The generated BrtExternSheet keeps the defined-name layout: entry 0 is the workbook, entry 1 is #REF!, entry 2 + n
/// is sheet n, and multi-sheet ranges are appended after the sheets in first-use order.
/// </remarks>
internal sealed class XlsbNewWorkbookFormulaContext : IXlsbFormulaEncodingContext {
    private readonly ExcelDocument _document;
    private readonly IReadOnlyList<ExcelSheet> _sheets;
    private readonly DefinedNames? _definedNames;
    private readonly Dictionary<ushort, ushort> _mapped = new();
    private readonly List<(int First, int Last)> _ranges = new();
    private LegacyXlsExternSheetTable? _links;
    private LegacyXlsFormulaNameIndex? _names;

    internal XlsbNewWorkbookFormulaContext(ExcelDocument document, IReadOnlyList<ExcelSheet> sheets, DefinedNames? definedNames) {
        _document = document;
        _sheets = sheets;
        _definedNames = definedNames;
    }

    /// <summary>Multi-sheet ranges referenced by encoded formulas, appended after the per-sheet entries.</summary>
    internal IReadOnlyList<(int First, int Last)> SheetRanges => _ranges;

    /// <summary>True once an encoded formula refers to the self-reference table.</summary>
    internal bool UsesExternalSheets => _mapped.Count != 0;

    public bool TryEncode(string text, string sheetName, out byte[] payload, out string? reason, bool arrayFormula = false, byte resultClass = 0) {
        _links ??= LegacyXlsExternSheetTable.Create(_document, _sheets);
        if (_links.SupportingLinkRecords.Count > 1) {
            throw new NotSupportedException("New external-workbook XLSB formulas require an external-link writer.");
        }
        _names ??= CreateNameIndex();
        int sheetIndex = -1;
        for (int index = 0; index < _sheets.Count; index++) {
            if (string.Equals(_sheets[index].Name, sheetName, StringComparison.Ordinal)) { sheetIndex = index; break; }
        }
        return XlsbFormulaEncoder.TryEncode(text, _names, sheetIndex, MapReference, out payload, out reason, arrayFormula, resultClass);
    }

    public XlsbWriteCell RebindSourceCell(XlsbWriteCell cell) => cell;

    public byte[] RebindSourceFormula(byte[] payload) => payload;

    // BrtName records follow the definedNames element order, so a name's one-based index is its position.
    private LegacyXlsFormulaNameIndex CreateNameIndex() {
        var global = new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase);
        var local = new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase);
        uint id = 0;
        foreach (DefinedName name in _definedNames?.Elements<DefinedName>() ?? Enumerable.Empty<DefinedName>()) {
            id++;
            string text = name.Name?.Value ?? string.Empty;
            if (name.LocalSheetId?.Value is uint scope) {
                local[LegacyXlsFormulaNameIndex.CreateLocalKey(checked((int)scope), text)] = id;
            } else {
                global[text] = id;
            }
        }
        return new LegacyXlsFormulaNameIndex(global, local, _links!);
    }

    private ushort MapReference(ushort generated) {
        if (_mapped.TryGetValue(generated, out ushort mapped)) return mapped;
        byte[] table = _links!.Payload;
        int offset = 2 + generated * 6;
        if (offset > table.Length - 6) throw new InvalidDataException("Invalid generated XLSB reference index.");
        int supportingLink = table[offset] | table[offset + 1] << 8;
        int first = table[offset + 2] | table[offset + 3] << 8;
        int last = table[offset + 4] | table[offset + 5] << 8;
        if (supportingLink != 0 || first >= _sheets.Count || last >= _sheets.Count) {
            throw new InvalidDataException("Invalid generated self-reference.");
        }
        if (first == last) {
            mapped = checked((ushort)(2 + first));
        } else {
            int range = _ranges.IndexOf((first, last));
            if (range < 0) { range = _ranges.Count; _ranges.Add((first, last)); }
            mapped = checked((ushort)(2 + _sheets.Count + range));
        }
        _mapped.Add(generated, mapped);
        return mapped;
    }
}

/// <summary>Encodes edited cell formulas for the native XLSB writers.</summary>
internal interface IXlsbFormulaEncodingContext {
    /// <remarks>resultClass is the operand class the formula result takes (ExcelNativeFormulaOperandTypes), or 0 for the
    /// cell default: value class, or array class for array formulas.</remarks>
    bool TryEncode(string text, string sheetName, out byte[] payload, out string? reason, bool arrayFormula = false, byte resultClass = 0);

    XlsbWriteCell RebindSourceCell(XlsbWriteCell cell);

    byte[] RebindSourceFormula(byte[] payload);
}
