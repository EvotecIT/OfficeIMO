namespace OfficeIMO.Epub;

/// <summary>A writable publication and the fidelity report from importing its manuscript.</summary>
public sealed class EpubManuscriptResult : IOfficeConversionResult<EpubPublication, EpubManuscriptReport> {
    internal EpubManuscriptResult(EpubPublication publication, IEnumerable<OfficeConversionFidelityDiagnostic> diagnostics) {
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
    /// <summary>Returns the publication only when every conversion stage completed successfully.</summary>
    public EpubPublication RequireValue() {
        if (!Succeeded) throw new InvalidOperationException("The manuscript import failed. Inspect its report before publishing.");
        return Publication;
    }
    /// <summary>Returns the publication only when every conversion stage completed without fidelity loss.</summary>
    public EpubPublication RequireNoLoss() { Report.RequireNoLoss(); return Publication; }
    /// <summary>Combines upstream conversion reports with this EPUB import without collapsing fidelity categories.</summary>
    public EpubManuscriptResult WithSourceReports(params IOfficeConversionReport[] reports) =>
        new EpubManuscriptResult(Publication, OfficeConversionFidelityDiagnostics.Flatten(reports).Concat(Report.FidelityDiagnostics));
}

/// <summary>Immutable category-preserving manuscript conversion evidence.</summary>
public sealed class EpubManuscriptReport : IOfficeConversionReport {
    internal EpubManuscriptReport(IEnumerable<OfficeConversionFidelityDiagnostic> diagnostics) =>
        FidelityDiagnostics = Array.AsReadOnly(diagnostics.ToArray());
    /// <summary>Diagnostics emitted by source conversion and EPUB packaging.</summary>
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics { get; }
    /// <summary>Whether every source conversion stage completed without a failure.</summary>
    public bool Succeeded => !FidelityDiagnostics.Any(item => item.LossKind == OfficeConversionLossKind.Failure);
    /// <summary>Whether any source content was approximated, omitted or failed.</summary>
    public bool HasLoss => FidelityDiagnostics.Any(item => item.LossKind != OfficeConversionLossKind.None);
    /// <summary>Rejects import results containing fidelity loss before they are published.</summary>
    public void RequireNoLoss() {
        if (HasLoss) throw new InvalidOperationException("The manuscript import reports content loss: " +
            string.Join("; ", FidelityDiagnostics.Where(item => item.LossKind != OfficeConversionLossKind.None).Select(item => item.Code)));
    }
}
