namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Only explicitly imported picture frames can exceed a paragraph's flow box.</summary>
    private static bool HasImportedOverflowPicture(IReadOnlyList<RichSeg> line) =>
        line.Any(segment => segment.InlineElement is PdfInlineImage { AllowFrameOverflow: true });

    /// <summary>
    /// Matches Word's paragraph-segment clip, including pictures on later exact-height lines.
    /// Word can paint over earlier lines in the same paragraph; their individual line tops are not clip boundaries.
    /// </summary>
    private static bool BeginImportedPictureClip(StringBuilder output, PdfInlineElement element,
        PdfLineSpacing? spacing, double x, double baseline, double paragraphTop, double pageWidth) {
        if (GetImportedPictureClip(element, spacing, x, baseline, paragraphTop, pageWidth) is not { } clip)
            return false;
        new ContentStreamBuilder(output).SaveState()
            .Rectangle(clip.X, clip.Bottom, clip.Width, clip.Height)
            .ClipPath().EndPath();
        return true;
    }

    /// <summary>Shares the authored picture-only paragraph clip with visible-paint diagnostics.</summary>
    private static (double X, double Bottom, double Width, double Height)? GetImportedPictureClip(
        PdfInlineElement element, PdfLineSpacing? spacing, double x, double baseline,
        double paragraphTop, double pageWidth) {
        if (element is not PdfInlineImage image || !(image.AllowFrameOverflow || image.ClipToParagraphFragment) ||
            spacing?.IsExact != true) return null;
        double bottom = baseline + element.BaselineOffset.ToPoints();
        double left = Math.Max(0D, x), right = Math.Min(pageWidth, x + element.Width.ToPoints());
        return (left, bottom, Math.Max(0D, right - left), Math.Max(0D, paragraphTop - bottom));
    }

    private sealed partial class LayoutContext {
        private bool CanOverflowImportedPictureLine(IReadOnlyList<RichSeg> line) =>
            !isRunningContent && HasImportedOverflowPicture(line);
    }
}
