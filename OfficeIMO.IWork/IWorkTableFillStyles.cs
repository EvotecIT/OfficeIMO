namespace OfficeIMO.IWork;

/// <summary>Supported native fill defaults for unbanded table regions. Null means absent or unresolved; an explicit no-fill remains a value.</summary>
public sealed class IWorkTableFillStyles {
    internal IWorkTableFillStyles(IWorkCellFill? body, IWorkCellFill? headerRow,
        IWorkCellFill? headerColumn, IWorkCellFill? footerRow) {
        Body = body; HeaderRow = headerRow; HeaderColumn = headerColumn; FooterRow = footerRow;
    }
    /// <summary>Gets body-cell defaults.</summary>
    public IWorkCellFill? Body { get; }
    /// <summary>Gets leading header-row defaults.</summary>
    public IWorkCellFill? HeaderRow { get; }
    /// <summary>Gets leading header-column defaults.</summary>
    public IWorkCellFill? HeaderColumn { get; }
    /// <summary>Gets trailing footer-row defaults.</summary>
    public IWorkCellFill? FooterRow { get; }
}
