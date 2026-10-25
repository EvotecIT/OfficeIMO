namespace OfficeIMO.Pdf;

/// <summary>
/// Holds layout warnings until a layout candidate is accepted. Rejected stabilization passes
/// cannot retain provisional clipping bounds or mutate the options' diagnostic deduplication caches.
/// </summary>
internal sealed class PdfLayoutDiagnosticBuffer : IDisposable {
    private static readonly System.Threading.AsyncLocal<PdfLayoutDiagnosticBuffer?> Active = new();
    private readonly PdfLayoutDiagnosticBuffer? previous;
    private readonly PdfConversionReport? owner;
    private readonly List<Action> pending = new();
    private bool disposed;

    internal PdfLayoutDiagnosticBuffer(PdfConversionReport? owner) {
        this.owner = owner;
        previous = Active.Value;
        Active.Value = this;
    }

    internal static bool TryCapture(PdfConversionReport owner, Action record) {
        // Options clones retain the report that owns their layout operation.
        // An independent nested write must complete its own report immediately.
        for (PdfLayoutDiagnosticBuffer? buffer = Active.Value; buffer != null; buffer = buffer.previous) {
            if (buffer.disposed || !ReferenceEquals(buffer.owner, owner)) continue;
            buffer.pending.Add(record);
            return true;
        }
        return false;
    }

    internal void Commit() {
#if NETSTANDARD2_0 || NET472
        if (disposed) throw new ObjectDisposedException(nameof(PdfLayoutDiagnosticBuffer));
#else
        ObjectDisposedException.ThrowIf(disposed, this);
#endif
        Active.Value = previous;
        disposed = true;
        try { foreach (Action record in pending) record(); }
        finally { pending.Clear(); }
    }

    public void Dispose() {
        if (disposed) return;
        Active.Value = previous;
        pending.Clear();
        disposed = true;
    }
}
