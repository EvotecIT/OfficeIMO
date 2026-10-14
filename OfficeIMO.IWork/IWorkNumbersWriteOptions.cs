namespace OfficeIMO.IWork;

/// <summary>Operation-level limits for template-free native Numbers creation.</summary>
public sealed class IWorkNumbersWriteOptions {
    /// <summary>Gets or sets the maximum sheet count.</summary>
    public int MaximumSheets { get; set; } = 1_024;
    /// <summary>Gets or sets the maximum combined table count.</summary>
    public int MaximumTables { get; set; } = 1_024;
    /// <summary>Gets or sets the maximum combined grid positions, including unset empty cells.</summary>
    public int MaximumTableCells { get; set; } = 65_536;
    /// <summary>Gets or sets the maximum combined UTF-16 length of sheet/table names and cell text.</summary>
    public int MaximumTextCharacters { get; set; } = 8_000_000;
    /// <summary>Gets or sets the maximum encoded package and intermediate archive size.</summary>
    public int MaximumPackageBytes { get; set; } = 64 * 1024 * 1024;

    internal IWorkNumbersWriteOptions Snapshot() {
        var copy = new IWorkNumbersWriteOptions { MaximumSheets = MaximumSheets, MaximumTables = MaximumTables,
            MaximumTableCells = MaximumTableCells, MaximumTextCharacters = MaximumTextCharacters,
            MaximumPackageBytes = MaximumPackageBytes };
        if (copy.MaximumSheets <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumSheets));
        if (copy.MaximumTables <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumTables));
        if (copy.MaximumTableCells <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumTableCells));
        if (copy.MaximumTextCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumTextCharacters));
        if (copy.MaximumPackageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumPackageBytes));
        return copy;
    }
}
