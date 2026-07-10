namespace OfficeIMO.Email;

internal sealed class MimeWriterState {
    private readonly HashSet<EmailDocument> _activeDocuments = new HashSet<EmailDocument>();

    internal MimeWriterState(EmailWriterOptions options, IList<EmailDiagnostic> diagnostics) {
        Options = options;
        Diagnostics = diagnostics;
    }

    internal EmailWriterOptions Options { get; }

    internal IList<EmailDiagnostic> Diagnostics { get; }

    internal void Enter(EmailDocument document, int depth) {
        if (depth > Options.MaxNestedMessageDepth) {
            throw new InvalidOperationException("The embedded-message write depth exceeds the configured maximum.");
        }
        if (!_activeDocuments.Add(document)) throw new InvalidOperationException("The embedded-message graph contains a cycle.");
    }

    internal void Exit(EmailDocument document) {
        _activeDocuments.Remove(document);
    }
}
