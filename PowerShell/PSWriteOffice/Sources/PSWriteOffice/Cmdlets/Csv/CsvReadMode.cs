namespace PSWriteOffice.Cmdlets.Csv;

/// <summary>Controls whether a CSV import streams rows or materializes the document first.</summary>
public enum CsvReadMode
{
    /// <summary>Read rows through the forward-only OfficeIMO data-reader surface.</summary>
    Stream,

    /// <summary>Load the complete editable CSV document before emitting rows.</summary>
    InMemory
}
