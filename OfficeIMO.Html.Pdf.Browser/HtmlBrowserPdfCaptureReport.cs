using System;
using HtmlTinkerX;
using OfficeIMO;

namespace OfficeIMO.Html.Pdf.Browser;

/// <summary>Browser-stage diagnostics attached to an OfficeIMO PDF conversion result.</summary>
public sealed class HtmlBrowserPdfCaptureReport : IOfficeConversionReport {
    /// <summary>Initializes a report from an immutable HtmlTinkerX capture result.</summary>
    public HtmlBrowserPdfCaptureReport(HtmlBrowserPdfDiagnostics diagnostics, bool tagged) {
        Diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
        Tagged = tagged;
    }

    /// <summary>Gets the HtmlTinkerX browser capture diagnostics.</summary>
    public HtmlBrowserPdfDiagnostics Diagnostics { get; }

    /// <summary>Gets whether Chromium was requested to generate a tagged PDF.</summary>
    public bool Tagged { get; }

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
