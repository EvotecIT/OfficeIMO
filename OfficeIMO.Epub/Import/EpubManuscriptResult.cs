namespace OfficeIMO.Epub;

/// <summary>A writable publication and the fidelity report from importing its manuscript.</summary>
public sealed class EpubManuscriptResult : IOfficeResult<EpubPublication> {
    internal EpubManuscriptResult(EpubPublication publication, IEnumerable<OfficeDiagnostic> diagnostics) {
        Publication = publication;
        Report = new EpubManuscriptReport(diagnostics);
    }
    /// <summary>Imported reflowable publication, ready for inspection and further editing.</summary>
    public EpubPublication Publication { get; }
    /// <summary>Imported publication through the shared conversion-result contract.</summary>
    public EpubPublication Value => Publication;
    /// <summary>Source content approximated or omitted during import.</summary>
    public EpubManuscriptReport Report { get; }
    /// <summary>Whether import completed without a failure diagnostic.</summary>
    public bool Succeeded => Report.Succeeded;
    /// <summary>Whether any conversion stage reports fidelity loss.</summary>
    public bool HasLoss => Report.HasLoss;

    /// <summary>Throws when the conversion report describes possible fidelity loss.</summary>
    public void ThrowIfLoss() => Report.ThrowIfLoss();
    /// <summary>Returns the publication only when every conversion stage completed successfully.</summary>
    public EpubPublication RequireValue() {
        if (!Succeeded) throw new InvalidOperationException("The manuscript import failed. Inspect its report before publishing.");
        return Publication;
    }
    /// <summary>Combines upstream conversion reports with this EPUB import without collapsing fidelity categories.</summary>
    public EpubManuscriptResult WithSourceReports(params OfficeConversionReport[] reports) =>
        new EpubManuscriptResult(Publication, OfficeConversionFidelityDiagnostics.Flatten(reports).Concat(Report.Diagnostics));
}

/// <summary>Immutable category-preserving manuscript conversion evidence.</summary>
public sealed class EpubManuscriptReport : OfficeConversionReport {
    internal const int MaximumDiagnostics = 10_000;
    internal EpubManuscriptReport(IEnumerable<OfficeDiagnostic> diagnostics) {
        var retained = diagnostics.Take(MaximumDiagnostics + 1).ToArray();
        if (retained.Length > MaximumDiagnostics) retained = retained.Take(MaximumDiagnostics - 1).Concat(new[] { LimitDiagnostic() }).ToArray();
        FidelityDiagnostics = Array.AsReadOnly(retained);
    }
    internal static OfficeDiagnostic LimitDiagnostic() => new global::OfficeIMO.OfficeDiagnostic(
        "EPUB_IMPORT_DIAGNOSTIC_LIMIT",
        global::OfficeIMO.OfficeDiagnosticSeverity.Warning,
        "The import exceeds the 10,000-finding review bound. Reduce the source's reported issues before publishing.",
        null,
        "OfficeIMO.Epub.Import",
        OfficeConversionLossKind.Failure);
    /// <summary>Diagnostics emitted by source conversion and EPUB packaging.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics { get; }
    /// <summary>Whether every source conversion stage completed without a failure.</summary>
    public bool Succeeded => !FidelityDiagnostics.Any(item => item.LossKind == OfficeConversionLossKind.Failure);
    /// <summary>Whether any source content was approximated, omitted or failed.</summary>
    public override bool HasLoss => FidelityDiagnostics.Any(item => item.LossKind != OfficeConversionLossKind.None);
    /// <summary>Rejects import results containing fidelity loss before they are published.</summary>
    protected override OfficeConversionException CreateLossException() => new OfficeConversionException("The manuscript import reports content loss: " +
            string.Join("; ", FidelityDiagnostics.Where(item => item.LossKind != OfficeConversionLossKind.None).Select(item => item.Code)), this);
}
