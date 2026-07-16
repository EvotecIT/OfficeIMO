namespace OfficeIMO.Reader.Notebook;

/// <summary>Controls bounded Jupyter Notebook projection.</summary>
public sealed class ReaderNotebookOptions {
    /// <summary>Gets or sets whether code cells are included. Default: true.</summary>
    public bool IncludeCodeCells { get; set; } = true;

    /// <summary>Gets or sets whether text-like code outputs are included. Default: true.</summary>
    public bool IncludeOutputs { get; set; } = true;

    /// <summary>Gets or sets the maximum cells inspected. Default: 10,000.</summary>
    public int MaxCells { get; set; } = 10_000;

    /// <summary>Gets or sets the maximum outputs inspected per code cell. Default: 100.</summary>
    public int MaxOutputsPerCell { get; set; } = 100;

    /// <summary>Gets or sets the maximum source characters retained per cell. Default: 1,000,000.</summary>
    public int MaxCellCharacters { get; set; } = 1_000_000;

    /// <summary>Gets or sets the maximum output characters retained per cell. Default: 100,000.</summary>
    public int MaxOutputCharactersPerCell { get; set; } = 100_000;

    internal ReaderNotebookOptions CloneValidated() {
        if (MaxCells < 1 || MaxCells > 100_000) throw new ArgumentOutOfRangeException(nameof(MaxCells));
        if (MaxOutputsPerCell < 0 || MaxOutputsPerCell > 10_000) throw new ArgumentOutOfRangeException(nameof(MaxOutputsPerCell));
        if (MaxCellCharacters < 1 || MaxCellCharacters > 16_000_000) throw new ArgumentOutOfRangeException(nameof(MaxCellCharacters));
        if (MaxOutputCharactersPerCell < 0 || MaxOutputCharactersPerCell > 16_000_000) {
            throw new ArgumentOutOfRangeException(nameof(MaxOutputCharactersPerCell));
        }
        return new ReaderNotebookOptions {
            IncludeCodeCells = IncludeCodeCells,
            IncludeOutputs = IncludeOutputs,
            MaxCells = MaxCells,
            MaxOutputsPerCell = MaxOutputsPerCell,
            MaxCellCharacters = MaxCellCharacters,
            MaxOutputCharactersPerCell = MaxOutputCharactersPerCell
        };
    }
}
