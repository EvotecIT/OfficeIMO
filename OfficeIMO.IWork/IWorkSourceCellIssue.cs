namespace OfficeIMO.IWork;

/// <summary>A materialized selected cell whose storage or value could not be decoded. This is source evidence, not a destination outcome or an omitted-cell count.</summary>
public sealed class IWorkSourceCellIssue {
    internal IWorkSourceCellIssue(IWorkTable table, IWorkTableCell cell) {
        TableIdentity = table.SourceIdentity;
        Row = cell.Row;
        Column = cell.Column;
        Message = cell.Error ?? "Cell contents could not be decoded.";
    }

    /// <summary>Gets the native table-info identity when available.</summary>
    public IWorkObjectIdentity? TableIdentity { get; }
    /// <summary>Gets the one-based source row established by selected cell storage.</summary>
    public int Row { get; }
    /// <summary>Gets the one-based source column established by selected cell storage.</summary>
    public int Column { get; }
    /// <summary>Gets the decoding failure retained on the projected cell.</summary>
    public string Message { get; }
}
