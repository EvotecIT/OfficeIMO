using OfficeIMO.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private static bool TryAddPreparedOutlinedText(
        PdfCore.PdfPageCanvas canvas, HtmlRenderText visual, OfficePreparedTextShaping prepared,
        RegisteredWebFonts webFonts, PdfCore.PdfConversionReport report, double frameWidth,
        bool asSpan, bool logicalTextOwned, CancellationToken cancellationToken,
        double baselineFontSize, bool suppressLink, bool preservePositionedFrame,
        ClipBounds? logicalClip) {
        cancellationToken.ThrowIfCancellationRequested();
        webFonts.OutlineBudget.ValidateTextLength(prepared.Text.Length);
        if (webFonts.OutlineBudget.IsPathLimitReached)
            throw new InvalidOperationException("HTML-to-PDF outlined text exceeded the configured path-command budget.");
        if (prepared.Program is not IOfficeBoundedFontProgram) {
            ReportPreparedShapingFallback(report, visual, "unbounded-selected-program");
            return false;
        }
        double advance = prepared.Measure(visual.Font.Size);
        if (double.IsNaN(advance) || double.IsInfinity(advance) || advance < 0D) {
            ReportPreparedShapingFallback(report, visual, "unrepresentable-selected-advance");
            return false;
        }
        bool painted = TryPaintOutlinedText(canvas, visual, webFonts, report, frameWidth,
            asSpan, logicalTextOwned, cancellationToken, baselineFontSize, suppressLink,
            preservePositionedFrame, logicalClip, new[] { new OutlinedFontRun(prepared, advance) });
        if (!painted) {
            ReportPreparedShapingFallback(report, visual, "selected-outlines-unavailable");
            return false;
        }
        bool alreadyReported = report.Warnings.Any(warning =>
            warning.Code == HtmlPdfDiagnosticCodes.FontProgramOutlined &&
            warning.Details.TryGetValue("PreparedShaping", out string? retained) && retained == "true" &&
            warning.Details.TryGetValue("Program", out string? program) &&
            string.Equals(program, prepared.Program.DisplayName ?? string.Empty, StringComparison.Ordinal));
        if (!alreadyReported) report.Add(new PdfCore.PdfConversionWarning(
            "OfficeIMO.Html.Pdf", HtmlPdfDiagnosticCodes.FontProgramOutlined,
            visual.Source ?? "html-text",
            "The selected font and retained glyph positions were painted as vector outlines with logical ActualText retained for extraction and accessibility.",
            PdfCore.PdfConversionWarningSeverity.Information,
            details: new Dictionary<string, string> {
                ["Family"] = visual.Font.FamilyName ?? string.Empty,
                ["Program"] = prepared.Program.DisplayName ?? string.Empty,
                ["PreparedShaping"] = "true",
                ["Representation"] = "vector-outlines-plus-actual-text"
            }));
        return true;
    }

    private static void ReportPreparedShapingFallback(
        PdfCore.PdfConversionReport report, HtmlRenderText visual, string reason) {
        if (report.Warnings.Any(warning =>
            warning.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported &&
            warning.Details.TryGetValue("Reason", out string? previous) && previous == reason &&
            string.Equals(warning.Source, visual.Source ?? "html-text", StringComparison.Ordinal))) return;
        report.Add(new PdfCore.PdfConversionWarning(
            "OfficeIMO.Html.Pdf", HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported,
            visual.Source ?? "html-text",
            "Retained glyph positions could not be represented by bounded outlines; PDF text uses an approximation that may differ from the admitted shaping result.",
            PdfCore.PdfConversionWarningSeverity.Warning, OfficeConversionLossKind.Approximation,
            details: new Dictionary<string, string> {
                ["FontFamily"] = visual.Font.FamilyName ?? string.Empty,
                ["Reason"] = reason,
                ["Fallback"] = "pdf-text"
            }));
    }
}
