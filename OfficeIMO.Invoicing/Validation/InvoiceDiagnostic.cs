namespace OfficeIMO.Invoicing;

/// <summary>Severity of an invoice diagnostic.</summary>
public enum InvoiceDiagnosticSeverity {
    /// <summary>Informational observation.</summary>
    Information,
    /// <summary>A non-blocking warning.</summary>
    Warning,
    /// <summary>An error that blocks valid output.</summary>
    Error
}

/// <summary>Structured validation or mapping diagnostic.</summary>
public sealed class InvoiceDiagnostic {
    /// <summary>Creates a diagnostic with a stable code and source location.</summary>
    public InvoiceDiagnostic(string code, string message, string location, InvoiceDiagnosticSeverity severity = InvoiceDiagnosticSeverity.Error) {
        Code = code; Message = message; Location = location; Severity = severity;
    }
    /// <summary>Stable rule or mapping code.</summary>
    public string Code { get; }
    /// <summary>Human-readable explanation.</summary>
    public string Message { get; }
    /// <summary>Model path or XML location.</summary>
    public string Location { get; }
    /// <summary>Diagnostic severity.</summary>
    public InvoiceDiagnosticSeverity Severity { get; }
}

/// <summary>Model validation result; this is not an authoritative schema or Schematron compliance certificate.</summary>
public sealed class InvoiceModelValidationResult {
    internal InvoiceModelValidationResult(List<InvoiceDiagnostic> diagnostics, InvoiceCalculation? calculation) {
        Diagnostics = diagnostics.AsReadOnly(); Calculation = calculation;
    }
    /// <summary>All model and declared-amount diagnostics.</summary>
    public IReadOnlyList<InvoiceDiagnostic> Diagnostics { get; }
    /// <summary>Calculated totals when arithmetic inputs permit calculation.</summary>
    public InvoiceCalculation? Calculation { get; }
    /// <summary>True when model validation produced no errors.</summary>
    public bool IsValid => Diagnostics.All(item => item.Severity != InvoiceDiagnosticSeverity.Error);
    /// <summary>Throws with the diagnostics when the model cannot be emitted safely.</summary>
    public void ThrowIfInvalid() {
        if (!IsValid) throw new InvalidDataException(string.Join(Environment.NewLine, Diagnostics
            .Where(item => item.Severity == InvoiceDiagnosticSeverity.Error).Select(item => item.Code + " " + item.Location + ": " + item.Message)));
    }
}
