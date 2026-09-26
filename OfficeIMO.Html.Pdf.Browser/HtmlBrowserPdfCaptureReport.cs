using System;
using System.Collections.Generic;
using HtmlTinkerX;
using OfficeIMO;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html.Pdf.Browser;

/// <summary>Browser-stage diagnostics attached to an OfficeIMO PDF conversion result.</summary>
public sealed class HtmlBrowserPdfCaptureReport : IOfficeConversionReport {
    /// <summary>Initializes a report from an immutable HtmlTinkerX capture result.</summary>
    public HtmlBrowserPdfCaptureReport(HtmlBrowserPdfDiagnostics diagnostics, bool tagged) {
        Diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
        Tagged = tagged;
        var fidelityDiagnostics = new List<OfficeConversionFidelityDiagnostic>();
        if (Diagnostics.BlockedRequestCount > 0) {
            fidelityDiagnostics.Add(new OfficeConversionFidelityDiagnostic(
                "HTML_BROWSER_BLOCKED_RESOURCES",
                Diagnostics.BlockedRequestCount + " browser resource request(s) were blocked.",
                OfficeConversionLossKind.Omission,
                "OfficeIMO.Html.Pdf.Browser"));
        }
        foreach (string warning in Diagnostics.Warnings) {
            fidelityDiagnostics.Add(new OfficeConversionFidelityDiagnostic(
                "HTML_BROWSER_WARNING",
                warning,
                OfficeConversionLossKind.Approximation,
                "OfficeIMO.Html.Pdf.Browser"));
        }
        fidelityDiagnostics.Add(new OfficeConversionFidelityDiagnostic(
            "HTML_BROWSER_NATIVE_TEXT_SHAPING",
            "Text layout was shaped by the browser engine rather than an OfficeIMO managed shaping provider.",
            OfficeConversionLossKind.None,
            "OfficeIMO.Html.Pdf.Browser"));
        FidelityDiagnostics = fidelityDiagnostics.AsReadOnly();
    }

    /// <summary>Gets the HtmlTinkerX browser capture diagnostics.</summary>
    public HtmlBrowserPdfDiagnostics Diagnostics { get; }

    /// <summary>Gets whether Chromium was requested to generate a tagged PDF.</summary>
    public bool Tagged { get; }

    /// <summary>Gets category-preserving browser-capture diagnostics.</summary>
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics { get; }

    /// <summary>Gets the shaping engine used by the browser capture path.</summary>
    public OfficeTextShapingBackend TextShapingBackend => OfficeTextShapingBackend.BrowserNative;

    /// <summary>
    /// Gets whether blocked resources or non-fatal browser warnings mean that captured content may be incomplete.
    /// </summary>
    public bool HasLoss => Diagnostics.BlockedRequestCount != 0 || Diagnostics.Warnings.Count != 0;

    /// <summary>Returns this report or throws when browser diagnostics indicate possible content loss.</summary>
    public HtmlBrowserPdfCaptureReport RequireNoLoss() {
        if (!HasLoss) return this;

        throw new InvalidOperationException(
            "Browser PDF capture may be incomplete: " +
            Diagnostics.BlockedRequestCount + " blocked request(s) and " +
            Diagnostics.Warnings.Count + " warning(s). Inspect the HtmlBrowserPdfCaptureReport.Diagnostics property for details.");
    }

    void IOfficeConversionReport.RequireNoLoss() => RequireNoLoss();
}
