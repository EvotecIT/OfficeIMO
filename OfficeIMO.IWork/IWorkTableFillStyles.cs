namespace OfficeIMO.IWork;

/// <summary>Supported native fill defaults for table regions and alternating body rows. Null means absent or unresolved; an explicit no-fill remains a value.</summary>
public sealed class IWorkTableFillStyles {
    internal IWorkTableFillStyles(IWorkCellFill? body, IWorkCellFill? headerRow,
        IWorkCellFill? headerColumn, IWorkCellFill? footerRow, IWorkCellFill? bandedBody = null) {
        Body = body; HeaderRow = headerRow; HeaderColumn = headerColumn; FooterRow = footerRow;
        BandedBody = bandedBody;
    }
    /// <summary>Gets the active fill for every second body row, counted after the header rows.
    /// Header columns and footer rows retain their region fills. Null means inactive, absent or unresolved.</summary>
    public IWorkCellFill? BandedBody { get; }
    /// <summary>Gets body-cell defaults.</summary>
    public IWorkCellFill? Body { get; }
    /// <summary>Gets leading header-row defaults.</summary>
    public IWorkCellFill? HeaderRow { get; }
    /// <summary>Gets leading header-column defaults.</summary>
    public IWorkCellFill? HeaderColumn { get; }
    /// <summary>Gets trailing footer-row defaults.</summary>
    public IWorkCellFill? FooterRow { get; }
}
