using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Utilities;

namespace OfficeIMO.Excel;

/// <summary>A manual page break attached to a worksheet.</summary>
public sealed class ExcelPageBreak {
    private readonly Break _element;

    internal ExcelPageBreak(ExcelSheet sheet, Break element, bool isRowBreak) {
        Sheet = sheet;
        _element = element;
        IsRowBreak = isRowBreak;
    }

    /// <summary>Gets the owning worksheet.</summary>
    public ExcelSheet Sheet { get; }

    /// <summary>Gets whether this break follows a row rather than a column.</summary>
    public bool IsRowBreak { get; }

    /// <summary>Gets the one-based row or column after which the page breaks.</summary>
    public int Index => Locking.ExecuteRead(Sheet.Document.EnsureLock(), () => {
        Sheet.EnsurePageBreakAttached(_element, IsRowBreak);
        return checked((int)_element.Id!.Value);
    });

    /// <summary>Removes this page break. A removed handle cannot affect a replacement at the same index.</summary>
    public void Remove(bool save = true) => Sheet.RemovePageBreak(_element, IsRowBreak, save);
}
