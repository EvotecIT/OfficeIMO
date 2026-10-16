using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Quantizes the selected face's ascent and descent independently for CSS
    /// inline boxes. The decoded program continues to expose its raw metrics.
    /// </summary>
    private static bool TryResolveUsedInlineFontBox(HtmlTextFaceMetrics? face, out InlineTextLineMetrics metrics) {
        metrics = default;
        if (!IsUsableTextFaceMetrics(face)) return false;
        double ascent = Math.Floor(face!.Value.BaselineOffset + 0.5D);
        double descent = Math.Floor(face.Value.Height - face.Value.BaselineOffset + 0.5D);
        metrics = new InlineTextLineMetrics(ascent + descent, ascent);
        return true;
    }

    /// <summary>
    /// Admits ordinary same-face text and empty struts at different sizes without
    /// broadening synthetic, mixed-face, displaced, bidi or atomic layout routes.
    /// </summary>
    private bool CanUseOrdinaryInlineFontBoxes(InlineLine line, HtmlRenderBoxStyle paragraphStyle) {
        if (!line.HasFlowContent || !IsOrdinaryInlineFontBoxStyle(paragraphStyle, paragraphStyle)) return false;
        IOfficeFontProgram? paragraphProgram = ResolveTextFaceProgram(string.Empty, paragraphStyle);
        if (paragraphProgram == null || !HasExactRegisteredFace(paragraphProgram, paragraphStyle)) return false;
        foreach (InlineSegment segment in line.Segments) {
            HtmlInlineRun run = segment.Run;
            if (!InlineLine.ContributesFlowContent(run)) continue;
            if (run.AtomicBlock != null || run.LeaderPattern != null || run.PaintOffsetX != 0D || run.PaintOffsetY != 0D
                || OfficeTextElements.ContainsRightToLeft(segment.Text)
                || !HasSameOrdinaryProgram(segment.Text, run.Style)) return false;
            foreach (HtmlRenderBoxStyle strut in run.InlineStrutStyles) {
                if (!HasSameOrdinaryProgram(string.Empty, strut)) return false;
            }
        }
        foreach (HtmlRenderBoxStyle strut in line.EmptyInlineStrutStyles) {
            if (!HasSameOrdinaryProgram(string.Empty, strut)) return false;
        }
        return true;

        bool HasSameOrdinaryProgram(string text, HtmlRenderBoxStyle style) {
            if (!IsOrdinaryInlineFontBoxStyle(style, paragraphStyle)) return false;
            IOfficeFontProgram? program = ResolveTextFaceProgram(text, style);
            return program != null && program.Fingerprint == paragraphProgram.Fingerprint
                && HasExactRegisteredFace(program, style);
        }
    }

    private bool HasExactRegisteredFace(IOfficeFontProgram program, HtmlRenderBoxStyle style) =>
        program is IOfficeFontBaselineMetrics && _fonts.Faces.Any(face =>
            face.Program.Fingerprint == program.Fingerprint && face.Descriptor.Equals(style.FontDescriptor));

    private static bool IsOrdinaryInlineFontBoxStyle(HtmlRenderBoxStyle style, HtmlRenderBoxStyle paragraphStyle) =>
        style.WritingMode == "horizontal-tb" && style.Direction == "ltr" && style.TableVerticalAlignment == "baseline"
        && style.BaselineScale == 1D && style.BaselineOffset == 0D && style.Font.Size > 0D
        && string.Equals(style.Font.FamilyName, paragraphStyle.Font.FamilyName, StringComparison.OrdinalIgnoreCase)
        && style.FontDescriptor.Equals(paragraphStyle.FontDescriptor);
}
