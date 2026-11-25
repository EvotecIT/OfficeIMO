namespace OfficeIMO.Notebook;

/// <summary>Sets hard resource limits for reading a Jupyter notebook.</summary>
public sealed class NotebookLoadOptions {
    /// <summary>Gets a fresh set of resource limits for external input.</summary>
    public static NotebookLoadOptions UntrustedDefaults => new();

    /// <summary>Gets or sets the largest accepted UTF-8 input, up to <see cref="Int32.MaxValue"/>. The default is 64 MiB.</summary>
    public long MaxInputBytes { get; set; } = 64L * 1024 * 1024;

    /// <summary>Gets or sets the maximum JSON nesting depth, up to 1000. The default is 128.</summary>
    public int MaxJsonDepth { get; set; } = 128;

    /// <summary>Gets or sets the maximum number of source cells. The default is 100,000.</summary>
    public int MaxCells { get; set; } = 100_000;

    /// <summary>Gets or sets the maximum number of output entries in one cell. The default is 10,000.</summary>
    public int MaxOutputsPerCell { get; set; } = 10_000;

    /// <summary>Gets or sets the maximum number of output entries in the notebook. The default is 100,000.</summary>
    public int MaxTotalOutputs { get; set; } = 100_000;

    internal NotebookLoadOptions Snapshot() {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(MaxInputBytes);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(MaxInputBytes, int.MaxValue);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(MaxJsonDepth);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(MaxJsonDepth, 1000);
        ArgumentOutOfRangeException.ThrowIfNegative(MaxCells);
        ArgumentOutOfRangeException.ThrowIfNegative(MaxOutputsPerCell);
        ArgumentOutOfRangeException.ThrowIfNegative(MaxTotalOutputs);
        return (NotebookLoadOptions)MemberwiseClone();
    }
}
