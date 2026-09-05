using OfficeIMO.PowerPoint;

namespace OfficeIMO.Markup.PowerPoint;

/// <summary>Immutable markup diagnostics and native PowerPoint preflight evidence.</summary>
public sealed class OfficeMarkupPowerPointConversionReport : IOfficeConversionReport {
    internal OfficeMarkupPowerPointConversionReport(
        IEnumerable<OfficeMarkupDiagnostic> diagnostics,
        PowerPointDeckPreflightReport preflightReport) {
        Markup = new OfficeMarkupConversionReport(diagnostics);
        Preflight = preflightReport ?? throw new ArgumentNullException(nameof(preflightReport));
    }

    /// <summary>Markup mapping diagnostics.</summary>
    public OfficeMarkupConversionReport Markup { get; }

    /// <summary>Native PowerPoint deck preflight report.</summary>
    public PowerPointDeckPreflightReport Preflight { get; }

    /// <summary>Markup mapping diagnostics in emission order.</summary>
    public IReadOnlyList<OfficeMarkupDiagnostic> Diagnostics => Markup.Diagnostics;

    /// <summary>Whether markup conversion completed without an error diagnostic.</summary>
    public bool Succeeded => Markup.Succeeded;

    /// <summary>Whether markup conversion reported possible content loss.</summary>
    public bool HasLoss => Markup.HasLoss;

    /// <summary>Throws when markup conversion failed.</summary>
    public void RequireSuccess() => Markup.RequireSuccess();

    /// <summary>Throws when markup conversion reported possible content loss.</summary>
    public void RequireNoLoss() => Markup.RequireNoLoss();
}

/// <summary>An editable PowerPoint presentation with mapping diagnostics and native preflight evidence.</summary>
public sealed class OfficeMarkupPowerPointConversionResult : OfficeConversionResult<PowerPointPresentation, OfficeMarkupPowerPointConversionReport> {
    internal OfficeMarkupPowerPointConversionResult(
        PowerPointPresentation value,
        IEnumerable<OfficeMarkupDiagnostic> diagnostics,
        PowerPointDeckPreflightReport preflightReport)
        : base(value, new OfficeMarkupPowerPointConversionReport(diagnostics, preflightReport)) { }

    /// <summary>Whether markup conversion completed without an error diagnostic.</summary>
    public override bool Succeeded => Report.Succeeded;

    /// <summary>Returns the presentation or throws the markup-specific exception when conversion failed.</summary>
    public override PowerPointPresentation RequireValue() {
        Report.RequireSuccess();
        return Value;
    }
}
