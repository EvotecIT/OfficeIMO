namespace OfficeIMO.Chm;

/// <summary>Operation-level CHM source and target fidelity evidence. It is not a layout or EPUB conformance certificate.</summary>
public sealed class ChmConversionReport : OfficeConversionReport {
    internal ChmConversionReport(IEnumerable<string> topicPaths, IEnumerable<OfficeDiagnostic> diagnostics) {
        TopicPaths = Array.AsReadOnly(topicPaths.ToArray());
        var retained = diagnostics.Take(10_001).ToArray();
        if (retained.Length > 10_000) retained = retained.Take(9_999).Concat(new[] {
            new global::OfficeIMO.OfficeDiagnostic("CHM_DIAGNOSTIC_LIMIT", global::OfficeIMO.OfficeDiagnosticSeverity.Warning, "The conversion exceeds the 10,000-finding review bound. Select fewer topics before publishing.", null, "OfficeIMO.Chm", OfficeConversionLossKind.Failure)
        }).ToArray();
        FidelityDiagnostics = Array.AsReadOnly(retained);
    }
    /// <summary>Selected topic paths in export order.</summary>
    public IReadOnlyList<string> TopicPaths { get; }
    /// <inheritdoc />
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics { get; }
    /// <inheritdoc />
    public override bool HasLoss => FidelityDiagnostics.Any(diagnostic => diagnostic.LossKind != OfficeConversionLossKind.None);
    /// <inheritdoc />
    protected override OfficeConversionException CreateLossException() => new OfficeConversionException("The CHM conversion reported possible fidelity loss.", this);
}

/// <summary>A CHM conversion value and its source/target fidelity report.</summary>
/// <typeparam name="T">Target artifact or document type.</typeparam>
public sealed class ChmConversionResult<T> : OfficeConversionResult<T> where T : class {
    /// <summary>Gets the typed conversion report.</summary>
    public new ChmConversionReport Report => (ChmConversionReport)base.Report;

    internal ChmConversionResult(T value, ChmConversionReport report) : base(value, report) { }
    /// <summary>Gets whether the conversion produced no failure diagnostic.</summary>
    public bool Succeeded => !Report.Diagnostics.Any(item => item.LossKind == OfficeConversionLossKind.Failure);
}
