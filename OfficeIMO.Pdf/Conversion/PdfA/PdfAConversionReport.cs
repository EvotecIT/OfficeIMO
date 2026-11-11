namespace OfficeIMO.Pdf;

/// <summary>Lists every change PDF/A conversion made, or every reason it refused the source.</summary>
public sealed class PdfAConversionReport : OfficeConversionReport {
    private const string Format = "PDF/A";

    internal PdfAConversionReport(
        PdfComplianceProfile profile,
        IReadOnlyList<PdfAConversionChange> changes,
        IReadOnlyList<PdfAConversionViolation> violations)
        : base("PDF", ToProfileName(profile)) {
        Profile = profile;
        Changes = changes;
        Violations = violations;
    }

    /// <summary>Gets the requested PDF/A profile.</summary>
    public PdfComplianceProfile Profile { get; }

    /// <summary>Gets whether a converted file was produced.</summary>
    public bool Succeeded => Violations.Count == 0;

    /// <summary>Gets every change, by object. Empty when conversion was refused, because nothing was written.</summary>
    public IReadOnlyList<PdfAConversionChange> Changes { get; }

    /// <summary>Gets every reason the source was refused, by object. Empty when conversion succeeded.</summary>
    public IReadOnlyList<PdfAConversionViolation> Violations { get; }

    /// <summary>Gets whether conversion was refused or removed or approximated source content.</summary>
    public override bool HasLoss => !Succeeded || Changes.Any(static change => change.LossKind != OfficeConversionLossKind.None);

    /// <inheritdoc />
    protected override IReadOnlyList<OfficeDiagnostic> GetDiagnostics() {
        var diagnostics = new List<OfficeDiagnostic>(Changes.Count + Violations.Count);
        foreach (PdfAConversionViolation violation in Violations) {
            diagnostics.Add(new OfficeDiagnostic("PdfA." + violation.Reason, OfficeDiagnosticSeverity.Error, violation.Message,
                violation.Location, Format, OfficeConversionLossKind.Failure));
        }
        foreach (PdfAConversionChange change in Changes) {
            diagnostics.Add(new OfficeDiagnostic("PdfA." + change.Kind,
                change.LossKind == OfficeConversionLossKind.None ? OfficeDiagnosticSeverity.Information : OfficeDiagnosticSeverity.Warning,
                change.Message, change.Location, Format, change.LossKind));
        }
        return diagnostics.AsReadOnly();
    }

    /// <inheritdoc />
    protected override OfficeConversionException CreateLossException() =>
        Succeeded
            ? base.CreateLossException()
            : new PdfAConversionException(this);

    internal static string ToProfileName(PdfComplianceProfile profile) => profile switch {
        PdfComplianceProfile.PdfA2B => "PDF/A-2b",
        PdfComplianceProfile.PdfA2U => "PDF/A-2u",
        PdfComplianceProfile.PdfA3B => "PDF/A-3b",
        PdfComplianceProfile.PdfA3U => "PDF/A-3u",
        PdfComplianceProfile.PdfA4 => "PDF/A-4",
        PdfComplianceProfile.PdfA4F => "PDF/A-4f",
        _ => profile.ToString()
    };
}

/// <summary>Raised when a PDF cannot be converted to PDF/A without changing its content. Nothing is written.</summary>
public sealed class PdfAConversionException : OfficeConversionException {
    internal PdfAConversionException(PdfAConversionReport report)
        : base(BuildMessage(report), report) {
        Report = report;
    }

    /// <summary>Gets the report listing every reason the source was refused.</summary>
    public new PdfAConversionReport Report { get; }

    private static string BuildMessage(PdfAConversionReport report) {
        var message = new StringBuilder("The PDF cannot be converted to ")
            .Append(PdfAConversionReport.ToProfileName(report.Profile))
            .Append(" without changing its content: ");
        int shown = Math.Min(report.Violations.Count, 3);
        for (int index = 0; index < shown; index++) {
            if (index > 0) message.Append("; ");
            message.Append(report.Violations[index].Location).Append(": ").Append(report.Violations[index].Message);
        }
        if (report.Violations.Count > shown) {
            message.Append("; and ").Append(report.Violations.Count - shown).Append(" more. Inspect Report.Violations.");
        }
        return message.ToString();
    }
}
