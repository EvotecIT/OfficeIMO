namespace OfficeIMO.IWork;

/// <summary>Operation-level limits for template-free native Numbers creation.</summary>
public sealed class IWorkNumbersWriteOptions {
    /// <summary>Gets or sets the maximum sheet count.</summary>
    public int MaxSheets { get; set; } = 1_024;
    /// <summary>Gets or sets the maximum combined table count.</summary>
    public int MaxTables { get; set; } = 1_024;
    /// <summary>Gets or sets the maximum combined grid positions, including unset empty cells.</summary>
    public int MaxTableCells { get; set; } = 65_536;
    /// <summary>Gets or sets the maximum combined UTF-16 length of sheet/table names and cell text.</summary>
    public int MaxTextCharacters { get; set; } = 8_000_000;
    /// <summary>Gets or sets the maximum encoded package and intermediate archive size.</summary>
    public int MaxPackageBytes { get; set; } = 64 * 1024 * 1024;

    internal IWorkNumbersWriteOptions Snapshot() {
        var copy = new IWorkNumbersWriteOptions { MaxSheets = MaxSheets, MaxTables = MaxTables,
            MaxTableCells = MaxTableCells, MaxTextCharacters = MaxTextCharacters,
            MaxPackageBytes = MaxPackageBytes };
        if (copy.MaxSheets <= 0) throw new ArgumentOutOfRangeException(nameof(MaxSheets));
        if (copy.MaxTables <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTables));
        if (copy.MaxTableCells <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTableCells));
        if (copy.MaxTextCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTextCharacters));
        if (copy.MaxPackageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPackageBytes));
        return copy;
    }
}
