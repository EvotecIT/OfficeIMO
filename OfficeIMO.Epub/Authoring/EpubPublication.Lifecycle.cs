namespace OfficeIMO.Epub;

public sealed partial class EpubPublication : IDisposable {
    private string? _filePath;

    /// <summary>Gets or sets the associated native save path without writing the publication.</summary>
    /// <remarks>File loads and successful file saves update this path. Stream operations leave it unchanged.</remarks>
    public string? FilePath {
        get => _filePath;
        set {
            if (value != null && string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("File path cannot be empty.", nameof(value));
            _filePath = value == null ? null : Path.GetFullPath(value);
        }
    }

    /// <summary>Completes the publication lifecycle without saving or closing caller-owned streams.</summary>
    /// <remarks>The publication retains its package in memory and owns no open file handles.</remarks>
    public void Dispose() { }
}
