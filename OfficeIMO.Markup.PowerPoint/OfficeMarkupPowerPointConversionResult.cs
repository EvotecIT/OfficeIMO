using OfficeIMO.PowerPoint;

namespace OfficeIMO.Markup.PowerPoint;

/// <summary>Immutable markup diagnostics and native PowerPoint preflight evidence.</summary>
public sealed class OfficeMarkupPowerPointConversionReport {
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
public sealed class OfficeMarkupPowerPointConversionResult {
    internal OfficeMarkupPowerPointConversionResult(
        PowerPointPresentation value,
        IEnumerable<OfficeMarkupDiagnostic> diagnostics,
        PowerPointDeckPreflightReport preflightReport) {
        Value = value ?? throw new ArgumentNullException(nameof(value));
        Report = new OfficeMarkupPowerPointConversionReport(diagnostics, preflightReport);
    }

    /// <summary>Editable presentation. The caller owns and disposes it.</summary>
    public PowerPointPresentation Value { get; }

    /// <summary>Conversion diagnostics and preflight evidence.</summary>
    public OfficeMarkupPowerPointConversionReport Report { get; }

    /// <summary>Whether markup conversion completed without an error diagnostic.</summary>
    public bool Succeeded => Report.Succeeded;

    /// <summary>Whether markup conversion reported possible content loss.</summary>
    public bool HasLoss => Report.HasLoss;

    /// <summary>Returns the presentation when markup conversion succeeded.</summary>
    public PowerPointPresentation RequireValue() {
        Report.RequireSuccess();
        return Value;
    }

    /// <summary>Returns the presentation only when markup conversion reported no possible loss.</summary>
    public PowerPointPresentation RequireNoLoss() {
        Report.RequireNoLoss();
        return Value;
    }
}
