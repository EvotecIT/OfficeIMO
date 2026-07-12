using System;

namespace OfficeIMO.Core;

/// <summary>Describes the severity of an OfficeIMO operation diagnostic.</summary>
public enum OfficeDiagnosticSeverity {
    /// <summary>Useful context that does not indicate a problem.</summary>
    Info,
    /// <summary>The operation completed, but the caller should review the result.</summary>
    Warning,
    /// <summary>The operation could not produce a valid result.</summary>
    Error
}

/// <summary>Describes how a diagnostic affected converted or generated content.</summary>
public enum OfficeDiagnosticImpact {
    /// <summary>No content was lost or changed.</summary>
    None,
    /// <summary>Content was represented using a simpler destination capability.</summary>
    Simplified,
    /// <summary>Content could not be represented and was omitted.</summary>
    Omitted,
    /// <summary>The operation or affected artifact failed.</summary>
    Failure
}

/// <summary>Provides one immutable diagnostic emitted by an OfficeIMO operation.</summary>
public sealed class OfficeDiagnostic {
    /// <summary>Creates a diagnostic with stable machine-readable and human-readable details.</summary>
    public OfficeDiagnostic(
        string code,
        OfficeDiagnosticSeverity severity,
        string message,
        OfficeDiagnosticImpact impact = OfficeDiagnosticImpact.None,
        string? component = null,
        string? feature = null,
        string? source = null,
        string? detail = null) {
        if (string.IsNullOrWhiteSpace(code)) throw new ArgumentException("A diagnostic code is required.", nameof(code));
        if (string.IsNullOrWhiteSpace(message)) throw new ArgumentException("A diagnostic message is required.", nameof(message));

        Code = code;
        Severity = severity;
        Message = message;
        Impact = impact;
        Component = component;
        Feature = feature;
        Source = source;
        Detail = detail;
    }

    /// <summary>Stable machine-readable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Diagnostic severity.</summary>
    public OfficeDiagnosticSeverity Severity { get; }

    /// <summary>Human-readable explanation.</summary>
    public string Message { get; }

    /// <summary>Effect on converted or generated content.</summary>
    public OfficeDiagnosticImpact Impact { get; }

    /// <summary>Component that emitted the diagnostic, when available.</summary>
    public string? Component { get; }

    /// <summary>Format feature associated with the diagnostic, when available.</summary>
    public string? Feature { get; }

    /// <summary>Source location or logical source identifier, when available.</summary>
    public string? Source { get; }

    /// <summary>Additional technical detail, when available.</summary>
    public string? Detail { get; }
}
