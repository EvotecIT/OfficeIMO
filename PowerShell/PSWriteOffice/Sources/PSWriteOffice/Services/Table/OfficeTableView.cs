namespace PSWriteOffice.Services.Table;

/// <summary>
/// Describes how tabular input should be projected before it is rendered.
/// </summary>
public enum OfficeTableView
{
    /// <summary>Keep input objects as rows and object properties as columns.</summary>
    Normal,

    /// <summary>Turn object properties into rows and input rows into columns.</summary>
    Transpose
}
