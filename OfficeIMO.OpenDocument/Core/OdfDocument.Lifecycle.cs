namespace OfficeIMO.OpenDocument;

public abstract partial class OdfDocument : IDisposable {
    /// <summary>Completes the document lifecycle without saving.</summary>
    /// <remarks>The document retains its package in memory and owns no open file handles.</remarks>
    public void Dispose() { }
}
