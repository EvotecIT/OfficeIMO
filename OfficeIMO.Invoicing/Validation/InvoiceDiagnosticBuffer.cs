namespace OfficeIMO.Invoicing;

/// <summary>Bounds reported diagnostics while retaining the severity of omitted findings.</summary>
internal sealed class InvoiceDiagnosticBuffer {
    private const int MaximumDetails = 999;
    private readonly List<InvoiceDiagnostic> _details = new List<InvoiceDiagnostic>();
    private int _omitted;
    private InvoiceDiagnosticSeverity _omittedSeverity;
    internal bool HasErrors { get; private set; }

    internal void Add(InvoiceDiagnostic diagnostic) {
        if (diagnostic.Severity == InvoiceDiagnosticSeverity.Error) HasErrors = true;
        if (_details.Count < MaximumDetails) _details.Add(diagnostic);
        else Omit(diagnostic.Severity);
    }

    internal void Add(string code, string message, string location, InvoiceDiagnosticSeverity severity = InvoiceDiagnosticSeverity.Error) {
        if (severity == InvoiceDiagnosticSeverity.Error) HasErrors = true;
        if (_details.Count < MaximumDetails) _details.Add(new InvoiceDiagnostic(code, message, location, severity));
        else Omit(severity);
    }

    internal void AddRange(IEnumerable<InvoiceDiagnostic> diagnostics) {
        foreach (InvoiceDiagnostic diagnostic in diagnostics) Add(diagnostic);
    }

    private void Omit(InvoiceDiagnosticSeverity severity) {
        if (_omitted < int.MaxValue) _omitted++;
        if (severity > _omittedSeverity) _omittedSeverity = severity;
    }

    internal List<InvoiceDiagnostic> ToList() {
        var result = new List<InvoiceDiagnostic>(_details);
        if (_omitted != 0) result.Add(new InvoiceDiagnostic("INV-DIAGNOSTICS-TRUNCATED",
            _omitted + " additional diagnostics were omitted; this summary retains their highest severity.", "Invoice", _omittedSeverity));
        return result;
    }
}
