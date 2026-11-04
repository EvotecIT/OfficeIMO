namespace PSWriteOffice.Cmdlets.Csv;

/// <summary>Result shape returned by Import-OfficeCsv.</summary>
public enum OfficeCsvOutput {
    /// <summary>One PSCustomObject per row.</summary>
    Object,
    /// <summary>One hashtable per row.</summary>
    Hashtable,
    /// <summary>One materialized DataTable.</summary>
    DataTable,
    /// <summary>A forward-only reader owned and disposed by the caller.</summary>
    DataReader,
    /// <summary>A live CsvDocument.</summary>
    Document
}
