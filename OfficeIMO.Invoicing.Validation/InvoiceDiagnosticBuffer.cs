namespace OfficeIMO.Invoicing.Validation;

/// <summary>Bounds reported diagnostics while retaining the severity of omitted findings.</summary>
internal sealed class InvoiceDiagnosticBuffer {
    private const int MaximumDetails = 999;
    private readonly List<InvoiceDiagnostic> _details = new List<InvoiceDiagnostic>();
    private int _omitted;
    private InvoiceDiagnosticSeverity _omittedSeverity;

    internal void Add(InvoiceDiagnostic diagnostic) {
        if (_details.Count < MaximumDetails) _details.Add(diagnostic);
        else {
            _omitted++;
            if (diagnostic.Severity > _omittedSeverity) _omittedSeverity = diagnostic.Severity;
        }
    }

    internal List<InvoiceDiagnostic> ToList() {
        var result = new List<InvoiceDiagnostic>(_details);
        if (_omitted != 0) result.Add(new InvoiceDiagnostic("INV-DIAGNOSTICS-TRUNCATED",
            _omitted + " additional diagnostics were omitted; this summary retains their highest severity.", "Invoice", _omittedSeverity));
        return result;
    }
}
