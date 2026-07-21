namespace OfficeIMO.Adf;

/// <summary>Severity of a lossy or noteworthy conversion decision.</summary>
public enum AdfConversionSeverity {
    Information,
    Warning,
    Error,
}

/// <summary>A conversion decision recorded for fidelity review.</summary>
public sealed class AdfConversionDiagnostic {
    internal AdfConversionDiagnostic(string code, string path, string message, AdfConversionSeverity severity) {
        Code = code;
        Path = path;
        Message = message;
        Severity = severity;
    }

    public string Code { get; }
    public string Path { get; }
    public string Message { get; }
    public AdfConversionSeverity Severity { get; }
}

/// <summary>Operation-scoped conversion evidence.</summary>
public sealed class AdfConversionReport {
    /// <summary>Creates a report from operation-scoped diagnostics.</summary>
    public AdfConversionReport(IEnumerable<AdfConversionDiagnostic> diagnostics) =>
        Diagnostics = diagnostics?.ToArray() ?? throw new ArgumentNullException(nameof(diagnostics));
    public IReadOnlyList<AdfConversionDiagnostic> Diagnostics { get; }
    public bool IsLossless => Diagnostics.All(item => item.Severity == AdfConversionSeverity.Information);
    public bool HasErrors => Diagnostics.Any(item => item.Severity == AdfConversionSeverity.Error);
}

/// <summary>A converted value and its fidelity report.</summary>
public sealed class AdfConversionResult<T> {
    internal AdfConversionResult(T value, IReadOnlyList<AdfConversionDiagnostic> diagnostics) {
        Value = value;
        Report = new AdfConversionReport(diagnostics);
    }

    public T Value { get; }
    public AdfConversionReport Report { get; }
}

/// <summary>Options for ADF projections.</summary>
public sealed class AdfConversionOptions {
    /// <summary>When true, visible placeholders are emitted for unsupported nodes with no projectable text.</summary>
    public bool EmitUnsupportedPlaceholders { get; set; }
}
