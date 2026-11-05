namespace OfficeIMO.Rtf;

public sealed partial class RtfDocument : IDisposable {
    /// <summary>Gets the associated destination set by path creation, loading or a successful file save.</summary>
    public string? FilePath { get; private set; }

    /// <summary>Creates an empty document associated with a path, without writing it.</summary>
    public static RtfDocument Create(string path) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        RtfDocument document = Create();
        document.FilePath = Path.GetFullPath(path);
        return document;
    }

    /// <summary>Saves to the associated path. In-memory documents must first supply a file destination.</summary>
    public void Save(RtfSaveOptions? options = null, Encoding? encoding = null) {
        Save(GetAssociatedPath(), options, encoding);
    }

    /// <summary>Asynchronously saves to the associated path without changing caller-owned streams.</summary>
    public Task SaveAsync(RtfSaveOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        return SaveAsync(GetAssociatedPath(), options, encoding, cancellationToken);
    }

    /// <summary>Releases the document without saving. The in-memory model owns no external resources.</summary>
    public void Dispose() {
    }

    private string GetAssociatedPath() => FilePath ?? throw new InvalidOperationException("The document has no associated path. Supply a destination to Save.");
}
