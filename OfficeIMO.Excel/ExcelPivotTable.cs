using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel;

/// <summary>An authored pivot table attached to a worksheet.</summary>
public sealed class ExcelPivotTable {
    private readonly PivotTablePart _part;
    internal ExcelPivotTable(ExcelSheet sheet, PivotTablePart part, string sourceRange, string destinationCell) {
        Sheet = sheet;
        _part = part;
        SourceRange = sourceRange;
        DestinationCell = destinationCell;
    }
    /// <summary>Gets the worksheet containing this pivot table.</summary>
    public ExcelSheet Sheet { get; }
    /// <summary>Gets the pivot table name.</summary>
    public string Name => _part.PivotTableDefinition!.Name!.Value!;
    /// <summary>Gets the source worksheet range.</summary>
    public string SourceRange { get; }
    /// <summary>Gets the top-left destination cell.</summary>
    public string DestinationCell { get; }
}
