using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private readonly struct SvgTextFontRun {
        internal SvgTextFontRun(string text, IOfficeFontProgram? program, double width, OfficeTextDirection direction) {
            Text = text;
            Program = program;
            Width = width;
            Direction = direction;
        }

        internal string Text { get; }
        internal IOfficeFontProgram? Program { get; }
        internal double Width { get; }
        internal OfficeTextDirection Direction { get; }
    }

    /// <summary>Uses the shared grapheme/face plan for both SVG advances and painted contours.</summary>
    private static IReadOnlyList<SvgTextFontRun> ResolveSvgTextFontRuns(string text, double fontSize,
        SvgPaintContext style, OfficeFontFaceCollection fonts, OfficeTextDirection direction,
        bool visualOrder, CancellationToken cancellationToken) {
        var result = new List<SvgTextFontRun>();
        if (visualOrder) {
            foreach (OfficeBidiTextRun bidi in OfficeBidiTextResolver.ResolveVisualRuns(text, direction, cancellationToken)) {
                AddRuns(bidi.Text, bidi.Direction, reverse: bidi.Direction == OfficeTextDirection.RightToLeft);
            }
        } else AddRuns(text, direction, reverse: false);
        return result;

        void AddRuns(string value, OfficeTextDirection runDirection, bool reverse) {
            cancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<OfficeFontFallbackRun> planned = fonts.PlanFallbackRuns(value, style.FontFamily, style.FontFace);
            for (int index = 0; index < planned.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                OfficeFontFallbackRun part = planned[reverse ? planned.Count - index - 1 : index];
                IOfficeFontProgram? program = fonts.ResolveForText(part.Text, part.FamilyRequest, style.FontFace, fontSize, out _)
                    ?? OfficeTrueTypeFont.TryLoadFontFamilyForText(part.FamilyRequest, style.FontFace, part.Text, out _)
                        ?.ForInstalledOpticalSize(fontSize, style.FontFace.Weight >= 600);
                double width = program?.Measure(part.Text, fontSize) ?? double.NaN;
                if (double.IsNaN(width) || double.IsInfinity(width) || width <= 0D) {
                    program = null;
                    width = EstimateSvgTextWidth(part.Text, fontSize);
                }
                result.Add(new SvgTextFontRun(part.Text, program, width, runDirection));
            }
        }
    }

    /// <summary>Combines fallback-face ink before applying paint, preserving one gradient/pattern bounding box.</summary>
    private static bool TryGetSvgTextRunContours(SvgTextRun run, OfficeFontFaceCollection fonts,
        SvgElementReferenceRegistry references, int pointAllowance, ref int unsupported, out List<List<OfficePoint>> contours) {
        contours = new List<List<OfficePoint>>();
        IReadOnlyList<SvgTextFontRun> parts = run.FontProgram != null
            ? new[] { new SvgTextFontRun(run.Text, run.FontProgram, run.Width / run.GlyphScale, run.TextDirection) }
            : ResolveSvgTextFontRuns(run.Text, run.FontSize, run.Style, fonts, run.TextDirection,
                visualOrder: true, references.CancellationToken);
        double x = run.X;
        foreach (SvgTextFontRun part in parts) {
            references.CancellationToken.ThrowIfCancellationRequested();
            // Unsupported outline providers must remain an explicit import loss, rather than
            // replacing a successfully resolved face with a different whole-string font.
            if (part.Program is not IOfficeBoundedFontProgram bounded) return false;
            double lineHeight = bounded.LineHeight(run.FontSize);
            double baselineOffset = bounded is IOfficeFontBaselineMetrics metrics
                ? metrics.BaselineOffset(run.FontSize) : lineHeight * 0.8D;
            if (double.IsNaN(lineHeight) || double.IsInfinity(lineHeight) || lineHeight <= 0D ||
                double.IsNaN(baselineOffset) || double.IsInfinity(baselineOffset)) return false;
            baselineOffset = Math.Max(0D, Math.Min(lineHeight, baselineOffset));
            OfficeManagedTextFallback shaped = OfficeManagedTextShaper.Resolve(
                part.Text, bounded, part.Direction, references.CancellationToken);
            if (shaped.Incomplete) unsupported++;
            string outlineText = shaped.Used ? shaped.Text : part.Text;
            List<List<OfficePoint>> ink = bounded is IOfficeCffBoundedFontProgram cff
                ? cff.GetTextContoursBounded(outlineText, x, run.Baseline - baselineOffset, run.FontSize,
                    pointAllowance, references.CancellationToken, references.CffOperationBudget)
                : bounded.GetTextContoursBounded(outlineText, x, run.Baseline - baselineOffset, run.FontSize,
                    pointAllowance, references.CancellationToken);
            foreach (List<OfficePoint> contour in ink) {
                if (contour.Count > pointAllowance) throw new InvalidOperationException("SVG text outlines exceed the shared point allowance.");
                pointAllowance -= contour.Count;
                contours.Add(contour);
            }
            x += part.Width;
        }
        return true;
    }
}
