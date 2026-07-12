using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Core;

/// <summary>Represents an OfficeIMO operation value together with immutable diagnostics.</summary>
public class OfficeOperationResult<T> {
    private readonly IReadOnlyList<OfficeDiagnostic> _diagnostics;

    /// <summary>Creates an operation result.</summary>
    public OfficeOperationResult(T value, IEnumerable<OfficeDiagnostic>? diagnostics = null) {
        Value = value;
        _diagnostics = Array.AsReadOnly((diagnostics ?? Enumerable.Empty<OfficeDiagnostic>()).ToArray());
    }

    /// <summary>Value produced by the operation.</summary>
    public T Value { get; }

    /// <summary>Diagnostics produced by the operation.</summary>
    public IReadOnlyList<OfficeDiagnostic> Diagnostics => _diagnostics;

    /// <summary>Whether the operation produced no error diagnostics.</summary>
    public bool Succeeded => !HasErrors;

    /// <summary>Whether at least one error diagnostic was produced.</summary>
    public bool HasErrors => _diagnostics.Any(static diagnostic => diagnostic.Severity == OfficeDiagnosticSeverity.Error);

    /// <summary>Returns the value or throws an <see cref="OfficeOperationException"/> when errors were reported.</summary>
    public T RequireValue() {
        if (HasErrors) throw new OfficeOperationException(_diagnostics);
        return Value;
    }
}

/// <summary>Represents a cross-format conversion value and its diagnostics.</summary>
public sealed class OfficeConversionResult<T> : OfficeOperationResult<T> {
    /// <summary>Creates a conversion result.</summary>
    public OfficeConversionResult(T value, IEnumerable<OfficeDiagnostic>? diagnostics = null)
        : base(value, diagnostics) {
    }

    /// <summary>Whether any content was simplified, omitted, or failed.</summary>
    public bool HasLoss => Diagnostics.Any(static diagnostic => diagnostic.Impact != OfficeDiagnosticImpact.None);
}

/// <summary>Thrown when a caller requires a value from an operation that reported errors.</summary>
public sealed class OfficeOperationException : InvalidOperationException {
    /// <summary>Creates an exception from operation diagnostics.</summary>
    public OfficeOperationException(IReadOnlyList<OfficeDiagnostic> diagnostics)
        : base(CreateMessage(diagnostics)) {
        Diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
    }

    /// <summary>Diagnostics that caused the operation to fail.</summary>
    public IReadOnlyList<OfficeDiagnostic> Diagnostics { get; }

    private static string CreateMessage(IReadOnlyList<OfficeDiagnostic>? diagnostics) {
        if (diagnostics == null) return "The OfficeIMO operation failed.";
        OfficeDiagnostic? error = diagnostics.FirstOrDefault(static diagnostic => diagnostic.Severity == OfficeDiagnosticSeverity.Error);
        return error == null
            ? "The OfficeIMO operation failed."
            : string.Concat("The OfficeIMO operation failed: ", error.Code, ": ", error.Message);
    }
}
