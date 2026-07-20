using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static System.Collections.Generic.IReadOnlyList<TextRun> StripRunLinksWhenCellLinked(System.Collections.Generic.IReadOnlyList<TextRun> runs, string? linkUri, string? linkDestinationName) {
        if (!HasCellLinkTarget(linkUri, linkDestinationName) || !runs.Any(run => run.LinkUri != null || run.LinkDestinationName != null)) {
            return runs;
        }

        var stripped = new System.Collections.Generic.List<TextRun>(runs.Count);
        foreach (TextRun run in runs) {
            stripped.Add(new TextRun(
                run.Text,
                run.Bold,
                run.Underline,
                run.Color,
                run.Italic,
                run.Strike,
                run.FontSize,
                run.Font,
                baseline: run.Baseline,
                tabLeader: run.TabLeader,
                tabAlignment: run.TabAlignment,
                backgroundColor: run.BackgroundColor,
                fontFamily: run.FontFamily));
        }

        return stripped;
    }

    private static bool HasCellLinkTarget(string? linkUri, string? linkDestinationName) =>
        !string.IsNullOrEmpty(linkUri) || !string.IsNullOrEmpty(linkDestinationName);

    private static double GetParagraphLeading(PdfParagraphStyle? style, double fontSize) {
        double multiplier = style?.LineHeight ?? 1.4;
        if (multiplier <= 0 || double.IsNaN(multiplier) || double.IsInfinity(multiplier)) {
            throw new ArgumentException("Paragraph line height must be a positive finite value.");
        }

        return fontSize * multiplier;
    }

    private static double GetParagraphSpacingBefore(PdfParagraphStyle? style) {
        double spacingBefore = style?.SpacingBefore ?? 0;
        if (spacingBefore < 0 || double.IsNaN(spacingBefore) || double.IsInfinity(spacingBefore)) {
            throw new ArgumentException("Paragraph spacing before must be a non-negative finite value.");
        }

        return spacingBefore;
    }

    private static double GetParagraphSpacingAfter(PdfParagraphStyle? style, double leading) {
        double spacingAfter = style?.SpacingAfter ?? leading * 0.3;
        if (spacingAfter < 0 || double.IsNaN(spacingAfter) || double.IsInfinity(spacingAfter)) {
            throw new ArgumentException("Paragraph spacing after must be a non-negative finite value.");
        }

        return spacingAfter;
    }

    private static double GetParagraphTabStopWidth(PdfParagraphStyle? style) {
        double tabStopWidth = style?.DefaultTabStopWidth ?? DefaultParagraphTabStopWidth;
        if (tabStopWidth <= 0 || double.IsNaN(tabStopWidth) || double.IsInfinity(tabStopWidth)) {
            throw new ArgumentException("Paragraph default tab stop width must be a positive finite value.");
        }

        return tabStopWidth;
    }

    private static PdfTabStop[]? GetParagraphTabStops(PdfParagraphStyle? style) =>
        style?.TabStops.Count > 0 ? style.TabStops.ToArray() : null;

    private static PdfHeadingStyle? ResolveHeadingStyle(HeadingBlock block, PdfOptions options) {
        return block.Style ?? options.DefaultHeadingStylesSnapshot?.GetSnapshot(block.Level);
    }

    private static double GetHeadingFontSize(HeadingBlock block, PdfHeadingStyle? style) {
        return style?.GetFontSize(block.Level) ?? PdfHeadingStyle.GetDefaultFontSize(block.Level);
    }

    private static double GetHeadingLeading(PdfHeadingStyle? style, double fontSize) {
        return style?.GetLeading(fontSize) ?? fontSize * 1.25D;
    }

    private static double GetHeadingSpacingAfter(PdfHeadingStyle? style, double leading) {
        return style?.GetSpacingAfter(leading) ?? leading * 0.25D;
    }

    private static bool GetHeadingBold(PdfHeadingStyle? style) {
        return style?.Bold ?? true;
    }

    private static PdfStandardFont GetHeadingFont(PdfOptions options, PdfHeadingStyle? style) {
        var normalFont = ChooseNormal(style?.Font ?? options.DefaultFont);
        return GetHeadingBold(style) ? ChooseBold(normalFont) : normalFont;
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<TextRun> CreateHeadingTextRuns(HeadingBlock heading, PdfHeadingStyle? style, PdfColor? color) =>
        System.Array.AsReadOnly(new[] {
            new TextRun(
                heading.Text,
                bold: GetHeadingBold(style),
                color: color,
                font: style?.Font,
                fontFamily: style?.FontFamily)
        });

    private static string GetHeadingFontResource(PdfHeadingStyle? style) {
        return GetHeadingBold(style) ? "F2" : "F1";
    }

    private static PdfListStyle? ResolveListStyle(BulletListBlock block, PdfOptions options) {
        return block.Style ?? options.DefaultListStyleSnapshot;
    }

    private static PdfListStyle? ResolveListStyle(NumberedListBlock block, PdfOptions options) {
        return block.Style ?? options.DefaultListStyleSnapshot;
    }

    private static double GetListFontSize(PdfListStyle? style, double defaultFontSize) {
        return style?.GetFontSize(defaultFontSize) ?? defaultFontSize;
    }

    private static double GetListLeading(PdfListStyle? style, double fontSize) {
        return style?.GetLeading(fontSize) ?? fontSize * 1.4D;
    }

    private static double GetListMarkerGap(PdfListStyle? style, double defaultGap) {
        return style?.GetMarkerGap(defaultGap) ?? defaultGap;
    }

    private static double GetListMarkerWidth(PdfListStyle? style, double estimatedWidth) {
        return Math.Max(estimatedWidth, style?.MarkerWidth ?? estimatedWidth);
    }

    private static double GetListMarkerFontSize(PdfListStyle? style, double listFontSize) {
        return style?.MarkerFontSize ?? listFontSize;
    }

    private static PdfStandardFont GetListMarkerFont(PdfListStyle? style, PdfStandardFont defaultFont) {
        PdfStandardFont normalFont = ChooseNormal(style?.MarkerFont ?? defaultFont);
        if (style?.MarkerBold == true && style.MarkerItalic) {
            return ChooseBoldItalic(normalFont);
        }

        if (style?.MarkerBold == true) {
            return ChooseBold(normalFont);
        }

        if (style?.MarkerItalic == true) {
            return ChooseItalic(normalFont);
        }

        return normalFont;
    }

    private static PdfNamedFontFace? GetListMarkerNamedFont(PdfListStyle? style, PdfOptions options) {
        if (style == null ||
            !options.TryResolveNamedFontFace(style.MarkerFontFamily, style.MarkerBold, style.MarkerItalic, out PdfNamedFontFace namedFont)) {
            return null;
        }

        return namedFont;
    }

    private static PdfAlign GetBulletMarkerAlign(PdfListStyle? style) {
        return style?.MarkerAlign ?? PdfAlign.Left;
    }

    private static PdfAlign GetNumberedMarkerAlign(PdfListStyle? style) {
        return style?.MarkerAlign ?? PdfAlign.Right;
    }

    private static double GetListItemSpacing(PdfListStyle? style, double leading) {
        return style?.GetItemSpacing(leading) ?? leading * 0.15D;
    }

    private static PanelStyle ResolvePanelStyle(PanelParagraphBlock block, PdfOptions options) {
        return block.Style ?? options.DefaultPanelStyleSnapshot ?? new PanelStyle();
    }

    private static PdfHorizontalRuleStyle ResolveHorizontalRuleStyle(HorizontalRuleBlock block, PdfOptions options) {
        return block.Style ?? options.DefaultHorizontalRuleStyleSnapshot ?? new PdfHorizontalRuleStyle();
    }

    private static PdfImageStyle ResolveImageStyle(ImageBlock block, PdfOptions options) {
        return block.Style ?? options.DefaultImageStyleSnapshot ?? new PdfImageStyle();
    }

    private static PdfDrawingStyle ResolveDrawingStyle(ShapeBlock block, PdfOptions options) {
        return block.Style ?? options.DefaultDrawingStyleSnapshot ?? new PdfDrawingStyle();
    }

    private static PdfDrawingStyle ResolveDrawingStyle(DrawingBlock block, PdfOptions options) {
        return block.Style ?? options.DefaultDrawingStyleSnapshot ?? new PdfDrawingStyle();
    }

    private static (double X, double Width, double FirstLineX, double FirstLineWidth) GetParagraphTextFrame(PdfParagraphStyle? style, double x, double width) {
        double leftIndent = style?.LeftIndent ?? 0;
        double rightIndent = style?.RightIndent ?? 0;
        double firstLineIndent = style?.FirstLineIndent ?? 0;
        if (leftIndent < 0 || double.IsNaN(leftIndent) || double.IsInfinity(leftIndent)) {
            throw new ArgumentException("Paragraph left indent must be a non-negative finite value.");
        }

        if (rightIndent < 0 || double.IsNaN(rightIndent) || double.IsInfinity(rightIndent)) {
            throw new ArgumentException("Paragraph right indent must be a non-negative finite value.");
        }

        if (double.IsNaN(firstLineIndent) || double.IsInfinity(firstLineIndent)) {
            throw new ArgumentException("Paragraph first line indent must be a finite value.");
        }

        if (leftIndent + firstLineIndent < 0) {
            throw new ArgumentException("Paragraph first line indent must not move text outside the left content frame.");
        }

        double textWidth = width - leftIndent - rightIndent;
        if (textWidth <= 0 || double.IsNaN(textWidth) || double.IsInfinity(textWidth)) {
            throw new ArgumentException("Paragraph left and right indents must leave a positive text width.");
        }

        double firstLineWidth = textWidth - firstLineIndent;
        if (firstLineWidth <= 0 || double.IsNaN(firstLineWidth) || double.IsInfinity(firstLineWidth)) {
            throw new ArgumentException("Paragraph first line indent must leave a positive text width.");
        }

        return (x + leftIndent, textWidth, x + leftIndent + firstLineIndent, firstLineWidth);
    }

    private static bool TryApplyWidowControl(PdfParagraphStyle? style, int totalLineCount, int startLineIndex, ref int take, ref double heightSum, System.Collections.Generic.List<double> lineHeights, bool canMoveToNextPage) {
        if (style == null || take <= 0) {
            return false;
        }

        int minimumOrphanLines = ResolveMinimumOrphanLines(style);
        int minimumWidowLines = ResolveMinimumWidowLines(style);
        if (minimumOrphanLines <= 1 && minimumWidowLines <= 1) {
            return false;
        }

        int remainingLineCount = totalLineCount - startLineIndex;
        int afterTake = remainingLineCount - take;
        if (afterTake <= 0) {
            return false;
        }

        if (take < minimumOrphanLines && canMoveToNextPage) {
            return true;
        }

        if (afterTake < minimumWidowLines) {
            int linesToMove = minimumWidowLines - afterTake;
            if (take - linesToMove >= minimumOrphanLines) {
                for (int index = 0; index < linesToMove; index++) {
                    take--;
                    heightSum -= lineHeights[startLineIndex + take];
                }
            } else if (canMoveToNextPage) {
                return true;
            }
        }

        return false;
    }

    private static int ResolveMinimumOrphanLines(PdfParagraphStyle style) =>
        style.MinimumOrphanLines > 0 ? style.MinimumOrphanLines : style.WidowControl ? 2 : 0;

    private static int ResolveMinimumWidowLines(PdfParagraphStyle style) =>
        style.MinimumWidowLines > 0 ? style.MinimumWidowLines : style.WidowControl ? 2 : 0;

}
