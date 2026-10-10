namespace OfficeIMO.Chm;

/// <summary>Operation-level CHM source and target fidelity evidence. It is not a layout or EPUB conformance certificate.</summary>
public sealed class ChmConversionReport : IOfficeConversionReport {
    internal ChmConversionReport(IEnumerable<string> topicPaths, IEnumerable<OfficeConversionFidelityDiagnostic> diagnostics) {
        TopicPaths = Array.AsReadOnly(topicPaths.ToArray());
        var retained = diagnostics.Take(10_001).ToArray();
        if (retained.Length > 10_000) retained = retained.Take(9_999).Concat(new[] {
            new OfficeConversionFidelityDiagnostic("CHM_DIAGNOSTIC_LIMIT", "The conversion exceeds the 10,000-finding review bound. Select fewer topics before publishing.", OfficeConversionLossKind.Failure, "OfficeIMO.Chm")
        }).ToArray();
        FidelityDiagnostics = Array.AsReadOnly(retained);
    }
    /// <summary>Selected topic paths in export order.</summary>
    public IReadOnlyList<string> TopicPaths { get; }
    /// <inheritdoc />
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics { get; }
    /// <inheritdoc />
    public bool HasLoss => FidelityDiagnostics.Any(diagnostic => diagnostic.LossKind != OfficeConversionLossKind.None);
    /// <inheritdoc />
    public void RequireNoLoss() {
        if (HasLoss) throw new OfficeConversionException("The CHM conversion reported possible fidelity loss.", this);
    }
}

/// <summary>A CHM conversion value and its source/target fidelity report.</summary>
/// <typeparam name="T">Target artifact or document type.</typeparam>
public sealed class ChmConversionResult<T> : OfficeConversionResult<T, ChmConversionReport> where T : class {
    internal ChmConversionResult(T value, ChmConversionReport report) : base(value, report) { }
    /// <inheritdoc />
    public override bool Succeeded => !Report.FidelityDiagnostics.Any(item => item.LossKind == OfficeConversionLossKind.Failure);
}
