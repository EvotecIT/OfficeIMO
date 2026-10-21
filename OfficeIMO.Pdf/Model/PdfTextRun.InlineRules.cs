namespace OfficeIMO.Pdf;

public sealed partial class PdfTextRun {
    /// <summary>Projects an inline visual while retaining its source run's font and line-box policy.</summary>
    internal static PdfTextRun Inline(PdfInlineElement element, PdfTextRun metricStyle) {
        Guard.NotNull(metricStyle, nameof(metricStyle));
        var run = new PdfTextRun(element, metricStyle);
        // A paragraph-mark strut can supply typography, but the projected
        // visual itself still owns ink.
        run.IsLineBoxStrut = false;
        return run;
    }
}
