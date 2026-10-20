namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Only explicitly imported picture frames can exceed a paragraph's flow box.</summary>
    private static bool HasImportedOverflowPicture(IReadOnlyList<RichSeg> line) =>
        line.Any(segment => segment.InlineElement is PdfInlineImage { AllowFrameOverflow: true });

    /// <summary>Exact-height paragraphs crop picture paint above their segment without clipping adjacent text.</summary>
    private static bool BeginImportedPictureClip(StringBuilder output, PdfInlineElement element,
        PdfLineSpacing? spacing, double x, double baseline, double paragraphTop, double pageWidth) {
        if (element is not PdfInlineImage { AllowFrameOverflow: true } || spacing?.IsExact != true) return false;
        double bottom = baseline + element.BaselineOffset;
        double left = Math.Max(0D, x), right = Math.Min(pageWidth, x + element.Width);
        new ContentStreamBuilder(output).SaveState()
            .Rectangle(left, bottom, Math.Max(0D, right - left), Math.Max(0D, paragraphTop - bottom))
            .ClipPath().EndPath();
        return true;
    }

    private sealed partial class LayoutContext {
        private bool CanOverflowImportedPictureLine(IReadOnlyList<RichSeg> line) =>
            !isRunningContent && HasImportedOverflowPicture(line);
    }
}
