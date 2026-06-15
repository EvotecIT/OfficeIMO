namespace OfficeIMO.Rtf;

/// <summary>
/// RTF table block composed of rows and cells.
/// </summary>
public sealed class RtfTable : IRtfBlock {
    private readonly List<RtfTableRow> _rows = new List<RtfTableRow>();

    /// <summary>Table rows.</summary>
    public IReadOnlyList<RtfTableRow> Rows => _rows.AsReadOnly();

    /// <summary>Adds a row to the table.</summary>
    public RtfTableRow AddRow() {
        var row = new RtfTableRow();
        _rows.Add(row);
        return row;
    }

    internal void AddParsedRow(RtfTableRow row) {
        _rows.Add(row ?? throw new ArgumentNullException(nameof(row)));
    }
}
