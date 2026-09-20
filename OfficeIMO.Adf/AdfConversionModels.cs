namespace OfficeIMO.Adf;

/// <summary>Severity of a lossy or noteworthy conversion decision.</summary>
public enum AdfConversionSeverity {
    /// <summary>Records a decision without marking the conversion as lossy.</summary>
    Information,
    /// <summary>Marks a conversion decision that requires fidelity review.</summary>
    Warning,
    /// <summary>Marks an error-level conversion diagnostic.</summary>
    Error,
}

/// <summary>A conversion decision recorded for fidelity review.</summary>
public sealed class AdfConversionDiagnostic {
    /// <summary>Creates a diagnostic to include in a conversion report.</summary>
    /// <param name="code">Non-empty identifier for the conversion decision.</param>
    /// <param name="path">Location associated with the decision; use <c>$</c> for the document root.</param>
    /// <param name="message">Non-empty explanation for a reader of the report.</param>
    /// <param name="severity">Severity used by the report's fidelity and error checks.</param>
    public AdfConversionDiagnostic(string code, string path, string message, AdfConversionSeverity severity) {
        if (string.IsNullOrWhiteSpace(code)) throw new ArgumentException("A diagnostic code is required.", nameof(code));
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (string.IsNullOrWhiteSpace(message)) throw new ArgumentException("A diagnostic message is required.", nameof(message));
        Code = code;
        Path = path;
        Message = message;
        Severity = severity;
        LossKind = ResolveLossKind(code, severity);
    }

    /// <summary>Creates a diagnostic with an explicit fidelity-loss category.</summary>
    /// <param name="code">Non-empty identifier for the conversion decision.</param>
    /// <param name="path">Location associated with the decision; use <c>$</c> for the document root.</param>
    /// <param name="message">Non-empty explanation for a reader of the report.</param>
    /// <param name="severity">Severity used by the report's fidelity and error checks.</param>
    /// <param name="lossKind">Exact fidelity-loss category.</param>
    public AdfConversionDiagnostic(
        string code,
        string path,
        string message,
        AdfConversionSeverity severity,
        OfficeConversionLossKind lossKind)
        : this(code, path, message, severity) => LossKind = lossKind;

    /// <summary>Gets the diagnostic identifier supplied by the converter or caller.</summary>
    public string Code { get; }

    /// <summary>Gets the location associated with the decision, typically a JSON-style ADF path.</summary>
    public string Path { get; }

    /// <summary>Gets the human-readable explanation of the conversion decision.</summary>
    public string Message { get; }

    /// <summary>Gets the severity used when evaluating fidelity and errors in a report.</summary>
    public AdfConversionSeverity Severity { get; }

    /// <summary>Gets the exact fidelity-loss category represented by this diagnostic.</summary>
    public OfficeConversionLossKind LossKind { get; }

    private static OfficeConversionLossKind ResolveLossKind(
        string code,
        AdfConversionSeverity severity) {
        if (severity == AdfConversionSeverity.Information) return OfficeConversionLossKind.None;
        if (severity == AdfConversionSeverity.Error) return OfficeConversionLossKind.Failure;
        return code switch {
            "MARKDOWN_UNSUPPORTED_BLOCK" or
            "MARKDOWN_UNSUPPORTED_INLINE" or
            "ADF_ROOT_PROPERTIES_DROPPED" or
            "ADF_EMPTY_PARAGRAPH_DROPPED" or
            "ADF_HEADING_PROPERTIES_DROPPED" or
            "ADF_CODE_PROPERTIES_DROPPED" or
            "ADF_UNSUPPORTED_NODE" or
            "ADF_TABLE_ATTRIBUTES_DROPPED" or
            "ADF_TABLE_CELL_ATTRIBUTES_DROPPED" or
            "ADF_LINK_ATTRIBUTES_DROPPED" => OfficeConversionLossKind.Omission,
            _ => OfficeConversionLossKind.Approximation
        };
    }
}

/// <summary>Operation-scoped conversion evidence.</summary>
public sealed class AdfConversionReport : IOfficeConversionReport {
    /// <summary>Creates a report by copying the supplied diagnostic sequence.</summary>
    /// <param name="diagnostics">Diagnostics to include in the report; an empty sequence is allowed.</param>
    public AdfConversionReport(IEnumerable<AdfConversionDiagnostic> diagnostics) {
        Diagnostics = diagnostics?.ToArray() ?? throw new ArgumentNullException(nameof(diagnostics));
        FidelityDiagnostics = Array.AsReadOnly(Diagnostics.Select(static diagnostic =>
            new OfficeConversionFidelityDiagnostic(
                diagnostic.Code,
                diagnostic.Message,
                diagnostic.LossKind,
                "OfficeIMO.Adf",
                diagnostic.Path)).ToArray());
    }
    /// <summary>Gets a report with no diagnostics, which is considered lossless.</summary>
    public static AdfConversionReport Empty { get; } = new AdfConversionReport(Array.Empty<AdfConversionDiagnostic>());
    /// <summary>Gets the diagnostics captured when this report was created.</summary>
    /// <remarks>The sequence is copied from the constructor input; the exposed collection is not an immutable snapshot.</remarks>
    public IReadOnlyList<AdfConversionDiagnostic> Diagnostics { get; }

    /// <summary>Gets category-preserving conversion diagnostics.</summary>
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics { get; }

    /// <summary>Gets whether every diagnostic is informational; an empty report is also lossless.</summary>
    public bool IsLossless => FidelityDiagnostics.All(item => item.LossKind == OfficeConversionLossKind.None);

    /// <summary>Gets whether at least one diagnostic has <see cref="AdfConversionSeverity.Error"/> severity.</summary>
    public bool HasErrors => Diagnostics.Any(item => item.Severity == AdfConversionSeverity.Error);
    /// <inheritdoc />
    public bool HasLoss => !IsLossless;

    /// <inheritdoc />
    public void RequireNoLoss() {
        if (HasLoss) {
            throw new InvalidOperationException("The ADF conversion reported possible content loss. Inspect Report.Diagnostics for details.");
        }
    }
}

/// <summary>A converted value and its fidelity report.</summary>
public sealed class AdfConversionResult<T> : OfficeConversionResult<T, AdfConversionReport> where T : class {
    internal AdfConversionResult(T value, IReadOnlyList<AdfConversionDiagnostic> diagnostics)
        : base(value, new AdfConversionReport(diagnostics)) { }
}

/// <summary>Options for ADF projections.</summary>
public sealed class AdfConversionOptions {
    /// <summary>When true, visible placeholders are emitted for unsupported nodes with no projectable text.</summary>
    public bool EmitUnsupportedPlaceholders { get; set; }
}
