using System;
using System.Collections.Generic;
using HtmlTinkerX;
using OfficeIMO;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html.Pdf.Browser;

/// <summary>Browser-stage diagnostics attached to an OfficeIMO PDF conversion result.</summary>
public sealed class HtmlBrowserPdfCaptureReport : OfficeConversionReport {
    /// <summary>Initializes a report from an immutable HtmlTinkerX capture result.</summary>
    public HtmlBrowserPdfCaptureReport(HtmlBrowserPdfDiagnostics diagnostics, bool tagged) {
        NativeDiagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
        Tagged = tagged;
        var fidelityDiagnostics = new List<OfficeDiagnostic>();
        if (NativeDiagnostics.BlockedRequestCount > 0) {
            fidelityDiagnostics.Add(new global::OfficeIMO.OfficeDiagnostic(
                "HTML_BROWSER_BLOCKED_RESOURCES",
                global::OfficeIMO.OfficeDiagnosticSeverity.Warning,
                NativeDiagnostics.BlockedRequestCount + " browser resource request(s) were blocked.",
                null,
                "OfficeIMO.Html.Pdf.Browser",
                OfficeConversionLossKind.Omission));
        }
        foreach (string warning in NativeDiagnostics.Warnings) {
            fidelityDiagnostics.Add(new global::OfficeIMO.OfficeDiagnostic(
                "HTML_BROWSER_WARNING",
                global::OfficeIMO.OfficeDiagnosticSeverity.Warning,
                warning,
                null,
                "OfficeIMO.Html.Pdf.Browser",
                OfficeConversionLossKind.Approximation));
        }
        fidelityDiagnostics.Add(new global::OfficeIMO.OfficeDiagnostic(
            "HTML_BROWSER_NATIVE_TEXT_SHAPING",
            global::OfficeIMO.OfficeDiagnosticSeverity.Information,
            "Text layout was shaped by the browser engine rather than an OfficeIMO managed shaping provider.",
            null,
            "OfficeIMO.Html.Pdf.Browser",
            OfficeConversionLossKind.None));
        FidelityDiagnostics = fidelityDiagnostics.AsReadOnly();
    }

    /// <summary>Gets the HtmlTinkerX browser capture diagnostics.</summary>
    public HtmlBrowserPdfDiagnostics NativeDiagnostics { get; }

    /// <summary>Gets whether Chromium was requested to generate a tagged PDF.</summary>
    public bool Tagged { get; }

    /// <summary>Gets category-preserving browser-capture diagnostics.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics { get; }

    /// <summary>Gets the shaping engine used by the browser capture path.</summary>
    public OfficeTextShapingBackend TextShapingBackend => OfficeTextShapingBackend.BrowserNative;

    /// <summary>
    /// Gets whether blocked resources or non-fatal browser warnings mean that captured content may be incomplete.
    /// </summary>
    public override bool HasLoss => NativeDiagnostics.BlockedRequestCount != 0 || NativeDiagnostics.Warnings.Count != 0;

    /// <summary>Creates the exception thrown when browser diagnostics indicate possible content loss.</summary>
    protected override OfficeConversionException CreateLossException() => new OfficeConversionException(
            "Browser PDF capture may be incomplete: " +
            NativeDiagnostics.BlockedRequestCount + " blocked request(s) and " +
            NativeDiagnostics.Warnings.Count + " warning(s). Inspect the HtmlBrowserPdfCaptureReport.NativeDiagnostics property for details.", this);
}
