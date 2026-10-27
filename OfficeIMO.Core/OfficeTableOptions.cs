namespace OfficeIMO;

/// <summary>Shared options for projecting typed rows into an editable native table.</summary>
public sealed class OfficeTableOptions {
    /// <summary>Includes a first row containing the selected member names. Defaults to true.</summary>
    public bool IncludeHeaders { get; set; } = true;

    /// <summary>Optional native table name or accessible title.</summary>
    public string? Name { get; set; }
}
