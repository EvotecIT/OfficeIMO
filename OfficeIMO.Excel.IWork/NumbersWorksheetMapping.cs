namespace OfficeIMO.Excel.IWork;

/// <summary>Connects one source Numbers sheet or table to its generated Excel worksheet.</summary>
public sealed class NumbersWorksheetMapping {
    internal NumbersWorksheetMapping(int sourceSheetIndex, string sourceSheetName,
        int? sourceTableIndex, string? sourceTableName, string requestedName, string destinationName) {
        SourceSheetIndex = sourceSheetIndex;
        SourceSheetName = sourceSheetName;
        SourceTableIndex = sourceTableIndex;
        SourceTableName = sourceTableName;
        RequestedName = requestedName;
        DestinationName = destinationName;
    }

    /// <summary>Gets the one-based source sheet index.</summary>
    public int SourceSheetIndex { get; }
    /// <summary>Gets the original source sheet name.</summary>
    public string SourceSheetName { get; }
    /// <summary>Gets the one-based source table index, or null for sheet-level text.</summary>
    public int? SourceTableIndex { get; }
    /// <summary>Gets the original source table name, or null for sheet-level text.</summary>
    public string? SourceTableName { get; }
    /// <summary>Gets the worksheet name requested before destination normalization.</summary>
    public string RequestedName { get; }
    /// <summary>Gets the unique worksheet name written to the destination.</summary>
    public string DestinationName { get; }
    /// <summary>Gets whether Excel naming restrictions required a rename.</summary>
    public bool WasRenamed => !string.Equals(RequestedName, DestinationName, StringComparison.Ordinal);
}
