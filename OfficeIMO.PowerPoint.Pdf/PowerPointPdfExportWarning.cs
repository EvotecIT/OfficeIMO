using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.PowerPoint.Pdf;

/// <summary>
/// Describes PowerPoint content that the first-party PDF adapter skipped or simplified.
/// </summary>
public sealed class PowerPointPdfExportWarning {
    /// <summary>Creates a PowerPoint-to-PDF export warning.</summary>
    public PowerPointPdfExportWarning(int slideNumber, string code, string message)
        : this(slideNumber, code, message, null) { }

    /// <summary>Creates a PowerPoint-to-PDF export warning with a reusable layout diagnostic.</summary>
    public PowerPointPdfExportWarning(int slideNumber, string code, string message, PdfCore.PdfLayoutDiagnostic? layoutDiagnostic)
        : this(slideNumber, code, message, layoutDiagnostic,
            layoutDiagnostic?.Kind == PdfCore.PdfLayoutDiagnosticKind.SkippedContent
                ? OfficeConversionLossKind.Omission : OfficeConversionLossKind.Approximation) { }

    /// <summary>Creates a warning that preserves the renderer's typed fidelity impact.</summary>
    public PowerPointPdfExportWarning(int slideNumber, string code, string message,
        PdfCore.PdfLayoutDiagnostic? layoutDiagnostic, OfficeConversionLossKind lossKind) {
        if (!Enum.IsDefined(typeof(OfficeConversionLossKind), lossKind))
            throw new ArgumentOutOfRangeException(nameof(lossKind));
        SlideNumber = slideNumber;
        Code = code ?? throw new ArgumentNullException(nameof(code));
        Message = message ?? throw new ArgumentNullException(nameof(message));
        LayoutDiagnostic = layoutDiagnostic;
        LossKind = lossKind;
    }

    /// <summary>One-based slide number where the warning originated.</summary>
    public int SlideNumber { get; }

    /// <summary>Stable warning code for callers that want to group export diagnostics.</summary>
    public string Code { get; }

    /// <summary>Human-readable warning message.</summary>
    public string Message { get; }

    /// <summary>Reusable layout or visual fidelity diagnostic, when the warning maps to one.</summary>
    public PdfCore.PdfLayoutDiagnostic? LayoutDiagnostic { get; }

    /// <summary>Typed fidelity impact, including content skipped by the PDF adapter.</summary>
    public OfficeConversionLossKind LossKind { get; }

    /// <summary>
    /// Converts this PowerPoint-specific warning to the shared PDF conversion warning contract.
    /// </summary>
    public PdfCore.PdfConversionWarning ToConversionWarning() {
        var details = new Dictionary<string, string> {
            ["SlideNumber"] = SlideNumber.ToString(System.Globalization.CultureInfo.InvariantCulture)
        };

        return new PdfCore.PdfConversionWarning(
            "OfficeIMO.PowerPoint.Pdf",
            Code,
            "Slide " + SlideNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
            Message,
            severity: PdfCore.PdfConversionWarningSeverity.Warning,
            lossKind: LossKind,
            layoutDiagnostic: LayoutDiagnostic,
            details: details);
    }
}
