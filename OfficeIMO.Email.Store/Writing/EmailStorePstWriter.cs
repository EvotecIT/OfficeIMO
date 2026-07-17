using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>
/// Incrementally creates a new dependency-free Unicode PST. Source PST/OST files are never modified.
/// </summary>
public sealed class EmailStorePstWriter : IDisposable {
    private readonly PstStoreWriterCore _core;
    private bool _completed;
    private bool _disposed;

    private EmailStorePstWriter(string destinationPath, EmailStorePstWriterOptions options) {
        _core = new PstStoreWriterCore(destinationPath, options);
    }

    private EmailStorePstWriter(PstStoreWriterCore core) {
        _core = core;
    }

    /// <summary>Creates a writer targeting a new PST path.</summary>
    public static EmailStorePstWriter Create(string destinationPath,
        EmailStorePstWriterOptions? options = null) {
        if (string.IsNullOrWhiteSpace(destinationPath)) {
            throw new ArgumentException("A destination path is required.", nameof(destinationPath));
        }
        return new EmailStorePstWriter(destinationPath, options ?? new EmailStorePstWriterOptions());
    }

    /// <summary>Resumes a writer from its last integrity-checked durable checkpoint.</summary>
    public static EmailStorePstWriter Resume(string checkpointPath,
        IProgress<EmailStorePstWriteProgress>? progress = null) {
        if (string.IsNullOrWhiteSpace(checkpointPath)) {
            throw new ArgumentException("A checkpoint path is required.", nameof(checkpointPath));
        }
        return new EmailStorePstWriter(PstStoreWriterCore.Resume(checkpointPath, progress));
    }

    /// <summary>
    /// Deletes a checkpoint and its exact writer-owned working files. A missing checkpoint is a no-op.
    /// </summary>
    public static void DeleteCheckpoint(string checkpointPath) =>
        PstStoreWriterCore.DeleteCheckpoint(checkpointPath);

    /// <summary>Configured durable checkpoint path, or null when resumability is disabled.</summary>
    public string? CheckpointPath {
        get { ThrowIfUnavailable(); return _core.CheckpointPath; }
    }

    /// <summary>Identifier of the mandatory Top of Personal Folders container.</summary>
    public string RootFolderId {
        get { ThrowIfUnavailable(); return _core.RootFolderId; }
    }

    internal string DeletedItemsFolderId {
        get { ThrowIfUnavailable(); return _core.DeletedItemsFolderId; }
    }

    internal string SearchRootFolderId {
        get { ThrowIfUnavailable(); return _core.SearchRootFolderId; }
    }

    /// <summary>Adds a folder. A null parent places it under <see cref="RootFolderId"/>.</summary>
    public string AddFolder(string name, string? parentFolderId = null,
        string? containerClass = null) {
        ThrowIfUnavailable();
        return _core.AddFolder(name, parentFolderId, containerClass);
    }

    /// <summary>Writes one message or typed Outlook item into a folder.</summary>
    public string AddItem(string folderId, EmailDocument document, bool isAssociated = false,
        CancellationToken cancellationToken = default) {
        ThrowIfUnavailable();
        if (document == null) throw new ArgumentNullException(nameof(document));
        return _core.AddItem(folderId, document, isAssociated, cancellationToken);
    }

    /// <summary>Flushes all working data and atomically commits a resumable integrity-checked checkpoint.</summary>
    public void Checkpoint() {
        ThrowIfUnavailable();
        _core.Checkpoint();
    }

    /// <summary>Marks this incomplete writer for cleanup instead of checkpoint retention on disposal.</summary>
    public void Abandon() {
        ThrowIfUnavailable();
        _core.Abandon();
    }

    /// <summary>Finalizes indexes and allocation maps, then atomically commits the PST.</summary>
    public EmailStorePstWriteReport Complete(CancellationToken cancellationToken = default) {
        ThrowIfUnavailable();
        EmailStorePstWriteReport report = _core.Complete(cancellationToken);
        _completed = true;
        return report;
    }

    /// <inheritdoc />
    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _core.Dispose();
    }

    private void ThrowIfUnavailable() {
        if (_disposed) throw new ObjectDisposedException(nameof(EmailStorePstWriter));
        if (_completed) throw new InvalidOperationException("The PST writer has already completed.");
    }
}
