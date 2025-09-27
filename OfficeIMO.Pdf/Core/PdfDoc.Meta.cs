namespace OfficeIMO.Pdf;

/// <summary>
/// Root PDF document container and fluent API for composing basic PDF files.
/// Mirrors the OfficeIMO.Markdown style (H1/H2/H3, paragraph) but targets PDF output.
/// </summary>
public sealed partial class PdfDoc {
    private readonly System.Collections.Generic.List<IPdfBlock> _blocks = new();
    private readonly PdfOptions _options;
    private readonly System.Collections.Generic.Stack<System.Collections.Generic.List<IPdfBlock>> _blockScopes;

    // Metadata
    private string? _title;
    private string? _author;
    private string? _subject;
    private string? _keywords;

    private PdfDoc(PdfOptions? options = null) {
        _options = options ?? new PdfOptions();
        _blockScopes = new System.Collections.Generic.Stack<System.Collections.Generic.List<IPdfBlock>>();
        _blockScopes.Push(_blocks);
    }

    /// <summary>
    /// Creates a new, empty PDF document with optional <paramref name="options"/>.
    /// </summary>
    /// <param name="options">Page size, margins and default font options. When null, sensible defaults are used.</param>
    /// <returns>New <see cref="PdfDoc"/> instance.</returns>
    public static PdfDoc Create(PdfOptions? options = null) => new PdfDoc(options);

    /// <summary>
    /// Sets PDF metadata. Only values provided are updated; missing parameters keep previous values.
    /// Pass an empty string to clear a previously assigned value.
    /// </summary>
    /// <param name="title">Document title metadata.</param>
    /// <param name="author">Document author metadata.</param>
    /// <param name="subject">Document subject metadata.</param>
    /// <param name="keywords">Document keywords metadata.</param>
    /// <returns>This <see cref="PdfDoc"/> for chaining.</returns>
    public PdfDoc Meta(string? title = null, string? author = null, string? subject = null, string? keywords = null) {
        if (title != null) {
            _title = title.Length == 0 ? null : title;
        }

        if (author != null) {
            _author = author.Length == 0 ? null : author;
        }

        if (subject != null) {
            _subject = subject.Length == 0 ? null : subject;
        }

        if (keywords != null) {
            _keywords = keywords.Length == 0 ? null : keywords;
        }
        return this;
    }

    // Internal getters for writer/compose
    internal System.Collections.Generic.IEnumerable<IPdfBlock> Blocks => _blocks;
    internal PdfOptions Options => _options;

    private System.Collections.Generic.List<IPdfBlock> CurrentBlocks => _blockScopes.Peek();

    private void AddBlock(IPdfBlock block) { CurrentBlocks.Add(block); }

    internal void AddPageBlock(PageBlock pageBlock) { Guard.NotNull(pageBlock, nameof(pageBlock)); AddBlock(pageBlock); }

    internal System.IDisposable PushBlockScope(System.Collections.Generic.List<IPdfBlock> blocks) {
        Guard.NotNull(blocks, nameof(blocks));
        _blockScopes.Push(blocks);
        return new Scope(this);
    }

    private void PopScope() { if (_blockScopes.Count > 1) _blockScopes.Pop(); }

    private sealed class Scope : System.IDisposable {
        private readonly PdfDoc _doc;
        private bool _disposed;
        public Scope(PdfDoc doc) { _doc = doc; }
        public void Dispose() {
            if (_disposed) return;
            _doc.PopScope();
            _disposed = true;
        }
    }
}

