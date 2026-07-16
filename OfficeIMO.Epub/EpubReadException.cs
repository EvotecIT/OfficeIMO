namespace OfficeIMO.Epub;

/// <summary>Represents a fatal EPUB package read failure with structured diagnostics.</summary>
public sealed class EpubReadException : IOException {
    internal EpubReadException(string message, IReadOnlyList<EpubDiagnostic> diagnostics, Exception? innerException = null)
        : base(message, innerException) {
        Diagnostics = diagnostics ?? Array.Empty<EpubDiagnostic>();
    }

    /// <summary>Structured diagnostics associated with the fatal read failure.</summary>
    public IReadOnlyList<EpubDiagnostic> Diagnostics { get; }
}
