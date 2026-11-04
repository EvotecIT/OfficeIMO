using OfficeIMO.PowerPoint;

namespace OfficeIMO.Markup.PowerPoint;

/// <summary>Immutable markup diagnostics and native PowerPoint preflight evidence.</summary>
public sealed class OfficeMarkupPowerPointConversionReport : OfficeConversionReport {
    internal OfficeMarkupPowerPointConversionReport(
        IEnumerable<OfficeMarkupDiagnostic> diagnostics,
        PowerPointDeckPreflightReport preflightReport) {
        Markup = new OfficeMarkupConversionReport(diagnostics);
        Preflight = preflightReport ?? throw new ArgumentNullException(nameof(preflightReport));
        FidelityDiagnostics = Array.AsReadOnly(Markup.Diagnostics
            .Concat(Preflight.Findings.Select(static finding => new global::OfficeIMO.OfficeDiagnostic(
                finding.Code,
                global::OfficeIMO.OfficeDiagnosticSeverity.Warning,
                finding.Message,
                GetPreflightLocation(finding),
                "OfficeIMO.Markup.PowerPoint.Preflight",
                GetPreflightLossKind(finding.Severity))))
            .ToArray());
    }

    /// <summary>Markup mapping diagnostics.</summary>
    public OfficeMarkupConversionReport Markup { get; }

    /// <summary>Native PowerPoint deck preflight report.</summary>
    public PowerPointDeckPreflightReport Preflight { get; }

    /// <summary>Markup mapping diagnostics in emission order.</summary>
    public IReadOnlyList<OfficeMarkupDiagnostic> NativeDiagnostics => Markup.NativeDiagnostics;

    /// <summary>Category-preserving markup and native PowerPoint preflight diagnostics.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics { get; }

    /// <summary>Whether markup conversion completed without an error diagnostic.</summary>
    public bool Succeeded => Markup.Succeeded;

    /// <summary>Whether markup conversion or native PowerPoint preflight reported possible content loss.</summary>
    public override bool HasLoss => FidelityDiagnostics.Any(static diagnostic =>
        diagnostic.LossKind != OfficeConversionLossKind.None);

    /// <summary>Raises markup loss first, then preflight findings, matching the order of the two evidence sources.</summary>
    protected override OfficeConversionException CreateLossException() {
        Markup.ThrowIfLoss();
        Preflight.ThrowIfFindings(PowerPointDeckPreflightSeverity.Warning);
        return base.CreateLossException();
    }

    /// <summary>Throws when markup conversion failed.</summary>
    public void RequireSuccess() => Markup.RequireSuccess();


    private static OfficeConversionLossKind GetPreflightLossKind(PowerPointDeckPreflightSeverity severity) => severity switch {
        PowerPointDeckPreflightSeverity.Error => OfficeConversionLossKind.Failure,
        PowerPointDeckPreflightSeverity.Warning => OfficeConversionLossKind.Approximation,
        _ => OfficeConversionLossKind.None
    };

    private static string GetPreflightLocation(PowerPointDeckPreflightFinding finding) {
        string location = "slide:" + (finding.SlideIndex + 1).ToString(CultureInfo.InvariantCulture);
        return finding.ShapeIndex.HasValue
            ? location + "/shape:" + (finding.ShapeIndex.Value + 1).ToString(CultureInfo.InvariantCulture)
            : location;
    }
}

/// <summary>An editable PowerPoint presentation with mapping diagnostics and native preflight evidence.</summary>
public sealed class OfficeMarkupPowerPointConversionResult : OfficeConversionResult<PowerPointDocument> {
    /// <summary>Gets the typed conversion report.</summary>
    public new OfficeMarkupPowerPointConversionReport Report => (OfficeMarkupPowerPointConversionReport)base.Report;

    internal OfficeMarkupPowerPointConversionResult(
        PowerPointDocument value,
        IEnumerable<OfficeMarkupDiagnostic> diagnostics,
        PowerPointDeckPreflightReport preflightReport)
        : base(value, new OfficeMarkupPowerPointConversionReport(diagnostics, preflightReport)) { }

    /// <summary>Whether markup conversion completed without an error diagnostic.</summary>
    public bool Succeeded => Report.Succeeded;

    /// <summary>Returns the presentation or throws the markup-specific exception when conversion failed.</summary>
    public PowerPointDocument RequireValue() {
        Report.RequireSuccess();
        return Value;
    }
}
