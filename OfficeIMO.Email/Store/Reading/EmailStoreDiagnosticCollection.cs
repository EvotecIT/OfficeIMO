using System.Collections;

namespace OfficeIMO.Email.Store;

/// <summary>Bounds repeated diagnostics emitted during selected reads of an open store.</summary>
internal sealed class EmailStoreDiagnosticCollection : IReadOnlyList<EmailStoreDiagnostic> {
    private const int MaximumDiagnostics = 10_000;
    private readonly List<EmailStoreDiagnostic> _items = new List<EmailStoreDiagnostic>();
    private readonly HashSet<string> _keys = new HashSet<string>(StringComparer.Ordinal);

    internal void Add(EmailStoreDiagnostic diagnostic) {
        if (_items.Count > MaximumDiagnostics) return;
        string key = string.Concat(diagnostic.Code.Length, ":", diagnostic.Code,
            diagnostic.Message.Length, ":", diagnostic.Message, ":", (int)diagnostic.Severity, ":", diagnostic.Location);
        if (!_keys.Add(key)) return;
        if (_items.Count == MaximumDiagnostics) {
            _keys.Clear();
            _items.Add(new EmailStoreDiagnostic("EMAIL_STORE_DIAGNOSTICS_TRUNCATED",
                "Further store diagnostics were omitted after the bounded diagnostic catalog was filled.",
                EmailStoreDiagnosticSeverity.Warning));
        } else _items.Add(diagnostic);
    }

    internal IReadOnlyList<EmailStoreDiagnostic> AsReadOnly() => this;
    public int Count => _items.Count;
    public EmailStoreDiagnostic this[int index] => _items[index];
    public IEnumerator<EmailStoreDiagnostic> GetEnumerator() => _items.GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
