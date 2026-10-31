using System;

namespace OfficeIMO;

/// <summary>Severity of a diagnostic. Severity and fidelity loss are independent.</summary>
public enum OfficeDiagnosticSeverity {
    /// <summary>An informational observation.</summary>
    Information,
    /// <summary>An observation requiring review.</summary>
    Warning,
    /// <summary>An error recorded by a completed stage; operation failures still throw.</summary>
    Error
}

/// <summary>An immutable finding with a stable code and optional source location and format.</summary>
public sealed class OfficeDiagnostic {
    /// <summary>Creates a finding. Set lossKind explicitly for approximation, omission or unassessed content; warnings alone do not imply loss.</summary>
    public OfficeDiagnostic(string code, OfficeDiagnosticSeverity severity, string message, string? location = null,
        string? format = null, OfficeConversionLossKind lossKind = OfficeConversionLossKind.None) {
        if (string.IsNullOrWhiteSpace(code)) throw new ArgumentException("A diagnostic code is required.", nameof(code));
        if (!Enum.IsDefined(typeof(OfficeDiagnosticSeverity), severity)) throw new ArgumentOutOfRangeException(nameof(severity));
        if (!Enum.IsDefined(typeof(OfficeConversionLossKind), lossKind)) throw new ArgumentOutOfRangeException(nameof(lossKind));
        Code = code;
        Severity = severity;
        Message = message ?? throw new ArgumentNullException(nameof(message));
        Location = location;
        Format = format;
        LossKind = lossKind;
    }
    /// <summary>Gets a stable catalog code, or an extension code supplied by a caller.</summary>
    public string Code { get; }
    /// <summary>Gets severity independently of fidelity loss.</summary>
    public OfficeDiagnosticSeverity Severity { get; }
    /// <summary>Gets a human-readable description.</summary>
    public string Message { get; }
    /// <summary>Gets a source part, page, sheet, span or other source location.</summary>
    public string? Location { get; }
    /// <summary>Gets the format or conversion stage that emitted the finding.</summary>
    public string? Format { get; }
    /// <summary>Gets the fidelity impact, including loss that does not raise warning severity.</summary>
    public OfficeConversionLossKind LossKind { get; }
}
