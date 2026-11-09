namespace OfficeIMO.IWork;

public sealed partial class IWorkSourceDocument : IDisposable {
    /// <summary>Gets the captured source file path; directory bundles and byte/stream loads have no file association.</summary>
    public string? FilePath { get; private set; }

    /// <summary>Completes this captured snapshot's lifetime. Captured data and views do not own input streams.</summary>
    public void Dispose() { }
}

public sealed partial class IWorkNativeEditSession : IDisposable {
    private string? _filePath;
    /// <summary>Gets the last successful output path, falling back to the captured source file.</summary>
    public string? FilePath { get => _filePath ?? Source.FilePath; private set => _filePath = value; }

    /// <summary>Saves to the associated path, preserving explicit overwrite and source-directory containment rules.</summary>
    public IWorkNativeSaveReport Save(bool overwrite = false, IWorkNativeSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        Save(FilePath ?? throw new InvalidOperationException("The edit session has no associated path. Save to a path first."), overwrite, options, cancellationToken);

    /// <summary>Completes the managed editing lifetime without disposing its independently owned source snapshot.</summary>
    public void Dispose() { }
}
