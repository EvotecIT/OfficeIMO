using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static class RtfPdfMapping {
    internal static PdfCore.PdfAlign ToPdfAlign(RtfTextAlignment alignment) {
        switch (alignment) {
            case RtfTextAlignment.Center:
                return PdfCore.PdfAlign.Center;
            case RtfTextAlignment.Right:
                return PdfCore.PdfAlign.Right;
            case RtfTextAlignment.Justify:
                return PdfCore.PdfAlign.Justify;
            default:
                return PdfCore.PdfAlign.Left;
        }
    }

    internal static PdfCore.PdfTextBaseline ToPdfBaseline(RtfVerticalPosition position) {
        switch (position) {
            case RtfVerticalPosition.Superscript:
                return PdfCore.PdfTextBaseline.Superscript;
            case RtfVerticalPosition.Subscript:
                return PdfCore.PdfTextBaseline.Subscript;
            default:
                return PdfCore.PdfTextBaseline.Normal;
        }
    }

    internal static PdfCore.PdfColor? ToPdfColor(RtfDocument document, int? oneBasedColorIndex) {
        if (!oneBasedColorIndex.HasValue || oneBasedColorIndex.Value <= 0) {
            return null;
        }

        int index = oneBasedColorIndex.Value - 1;
        if (index < 0 || index >= document.Colors.Count) {
            return null;
        }

        RtfColor color = document.Colors[index];
        return PdfCore.PdfColor.FromRgb(color.Red, color.Green, color.Blue);
    }

    internal static PdfCore.PdfStandardFont? ToPdfFont(RtfDocument document, int? fontId, bool bold, bool italic) {
        if (!fontId.HasValue) {
            return null;
        }

        RtfFont? font = document.Fonts.FirstOrDefault(item => item.Id == fontId.Value);
        if (font == null) {
            return null;
        }

        return PdfCore.PdfStandardFontMapper.TryMapFontFamily(font.Name, bold, italic, out PdfCore.PdfStandardFont mapped)
            ? mapped
            : (PdfCore.PdfStandardFont?)null;
    }

    internal static PdfCore.PdfPageNumberStyle ToPdfPageNumberStyle(RtfPageNumberFormat format) {
        switch (format) {
            case RtfPageNumberFormat.UpperRoman:
                return PdfCore.PdfPageNumberStyle.UpperRoman;
            case RtfPageNumberFormat.LowerRoman:
                return PdfCore.PdfPageNumberStyle.LowerRoman;
            case RtfPageNumberFormat.UpperLetter:
                return PdfCore.PdfPageNumberStyle.UpperLetter;
            case RtfPageNumberFormat.LowerLetter:
                return PdfCore.PdfPageNumberStyle.LowerLetter;
            default:
                return PdfCore.PdfPageNumberStyle.Arabic;
        }
    }

    internal static PdfCore.PdfParagraphStyle? ToPdfParagraphStyle(RtfDocument document, RtfParagraph paragraph) {
        RtfStyle? style = paragraph.StyleId.HasValue
            ? document.Styles.FirstOrDefault(item => item.Id == paragraph.StyleId.Value && item.Kind == RtfStyleKind.Paragraph)
            : null;

        int? leftIndent = paragraph.LeftIndentTwips ?? style?.LeftIndentTwips;
        int? rightIndent = paragraph.RightIndentTwips ?? style?.RightIndentTwips;
        int? firstLineIndent = paragraph.FirstLineIndentTwips ?? style?.FirstLineIndentTwips;
        int? spaceBefore = paragraph.SpaceBeforeTwips ?? style?.SpaceBeforeTwips;
        int? spaceAfter = paragraph.SpaceAfterTwips ?? style?.SpaceAfterTwips;
        bool? spaceBeforeAuto = paragraph.SpaceBeforeAuto ?? style?.SpaceBeforeAuto;
        bool? spaceAfterAuto = paragraph.SpaceAfterAuto ?? style?.SpaceAfterAuto;
        int? lineSpacing = paragraph.LineSpacingTwips ?? style?.LineSpacingTwips;
        bool? lineSpacingMultiple = paragraph.LineSpacingMultiple ?? style?.LineSpacingMultiple;
        bool keepTogether = paragraph.KeepLinesTogether || style?.KeepLinesTogether == true;
        bool keepWithNext = paragraph.KeepWithNext || style?.KeepWithNext == true;
        bool widowControl = paragraph.WidowControl ?? style?.WidowControl ?? false;
        int? defaultTabWidth = document.Settings.DefaultTabWidthTwips;
        IReadOnlyList<RtfTabStop> tabStops = paragraph.TabStops.Count > 0
            ? paragraph.TabStops
            : style?.TabStops ?? Array.Empty<RtfTabStop>();
        bool hasTabStops = tabStops.Count > 0;

        if (!HasParagraphLayout(leftIndent, rightIndent, firstLineIndent, spaceBefore, spaceAfter, spaceBeforeAuto, spaceAfterAuto, lineSpacing, keepTogether, keepWithNext, widowControl, defaultTabWidth, hasTabStops)) {
            return null;
        }

        PdfCore.PdfParagraphStyle pdfStyle = new PdfCore.PdfParagraphStyle();
        double leftIndentPoints = ToNonNegativePoints(leftIndent);
        pdfStyle.LeftIndent = leftIndentPoints;
        pdfStyle.RightIndent = ToNonNegativePoints(rightIndent);
        pdfStyle.FirstLineIndent = ToSafeFirstLineIndent(firstLineIndent, leftIndentPoints);

        if (spaceBefore.HasValue && spaceBefore.Value >= 0) {
            pdfStyle.SpacingBefore = TwipsToPoints(spaceBefore.Value);
        } else if (spaceBeforeAuto == true) {
            pdfStyle.SpacingBefore = 0D;
        }

        if (spaceAfter.HasValue && spaceAfter.Value >= 0) {
            pdfStyle.SpacingAfter = TwipsToPoints(spaceAfter.Value);
        } else if (spaceAfterAuto == true) {
            pdfStyle.SpacingAfter = null;
        }

        double? lineHeight = ToPdfLineHeight(lineSpacing, lineSpacingMultiple, GetParagraphBaseFontSize(paragraph));
        if (lineHeight.HasValue) {
            pdfStyle.LineHeight = lineHeight.Value;
        }

        if (defaultTabWidth.HasValue && defaultTabWidth.Value > 0) {
            pdfStyle.DefaultTabStopWidth = TwipsToPoints(defaultTabWidth.Value);
        }

        foreach (RtfTabStop tabStop in tabStops) {
            pdfStyle.AddTabStop(TwipsToPoints(tabStop.PositionTwips), ToPdfTabAlignment(tabStop.Alignment), ToPdfTabLeader(tabStop.Leader));
        }

        pdfStyle.KeepTogether = keepTogether;
        pdfStyle.KeepWithNext = keepWithNext;
        pdfStyle.WidowControl = widowControl;
        return pdfStyle;
    }

    internal static bool HasPageBreakBefore(RtfDocument document, RtfParagraph paragraph) {
        if (paragraph.PageBreakBefore) {
            return true;
        }

        return paragraph.StyleId.HasValue &&
               document.Styles.FirstOrDefault(item => item.Id == paragraph.StyleId.Value && item.Kind == RtfStyleKind.Paragraph)?.PageBreakBefore == true;
    }

    internal static PdfCore.PdfPageBorder? ToPdfPageBorder(RtfDocument document, RtfPageBorders borders) {
        RtfPageBorder? source = GetFirstRenderablePageBorder(borders);
        if (source == null) {
            return null;
        }

        PdfCore.PdfPageBorder border = new PdfCore.PdfPageBorder {
            DashStyle = ToPdfDashStyle(source.Style),
            Width = source.Width.HasValue && source.Width.Value > 0 ? source.Width.Value / 8D : 1D,
            Inset = source.Space.HasValue && source.Space.Value >= 0 ? source.Space.Value : 36D
        };

        PdfCore.PdfColor? color = ToPdfColor(document, source.ColorIndex);
        if (color.HasValue) {
            border.Color = color.Value;
        }

        return border;
    }

    private static RtfPageBorder? GetFirstRenderablePageBorder(RtfPageBorders borders) {
        if (IsRenderablePageBorder(borders.Top)) return borders.Top;
        if (IsRenderablePageBorder(borders.Bottom)) return borders.Bottom;
        if (IsRenderablePageBorder(borders.Left)) return borders.Left;
        if (IsRenderablePageBorder(borders.Right)) return borders.Right;
        return null;
    }

    private static bool IsRenderablePageBorder(RtfPageBorder border) => border.Style != RtfPageBorderStyle.None;

    private static OfficeIMO.Drawing.OfficeStrokeDashStyle ToPdfDashStyle(RtfPageBorderStyle style) {
        switch (style) {
            case RtfPageBorderStyle.Dashed:
                return OfficeIMO.Drawing.OfficeStrokeDashStyle.Dash;
            case RtfPageBorderStyle.Dotted:
                return OfficeIMO.Drawing.OfficeStrokeDashStyle.Dot;
            default:
                return OfficeIMO.Drawing.OfficeStrokeDashStyle.Solid;
        }
    }

    private static bool HasParagraphLayout(
        int? leftIndent,
        int? rightIndent,
        int? firstLineIndent,
        int? spaceBefore,
        int? spaceAfter,
        bool? spaceBeforeAuto,
        bool? spaceAfterAuto,
        int? lineSpacing,
        bool keepTogether,
        bool keepWithNext,
        bool widowControl,
        int? defaultTabWidth,
        bool hasTabStops) =>
        leftIndent.HasValue ||
        rightIndent.HasValue ||
        firstLineIndent.HasValue ||
        spaceBefore.HasValue ||
        spaceAfter.HasValue ||
        spaceBeforeAuto.HasValue ||
        spaceAfterAuto.HasValue ||
        lineSpacing.HasValue ||
        keepTogether ||
        keepWithNext ||
        widowControl ||
        (defaultTabWidth.HasValue && defaultTabWidth.Value > 0) ||
        hasTabStops;

    private static PdfCore.PdfTabAlignment ToPdfTabAlignment(RtfTabAlignment alignment) {
        switch (alignment) {
            case RtfTabAlignment.Center:
                return PdfCore.PdfTabAlignment.Center;
            case RtfTabAlignment.Right:
                return PdfCore.PdfTabAlignment.Right;
            case RtfTabAlignment.Decimal:
                return PdfCore.PdfTabAlignment.DecimalSeparator;
            default:
                return PdfCore.PdfTabAlignment.Left;
        }
    }

    private static PdfCore.PdfTabLeaderStyle ToPdfTabLeader(RtfTabLeader leader) {
        switch (leader) {
            case RtfTabLeader.Dots:
            case RtfTabLeader.MiddleDots:
                return PdfCore.PdfTabLeaderStyle.Dots;
            case RtfTabLeader.Hyphen:
                return PdfCore.PdfTabLeaderStyle.Hyphens;
            case RtfTabLeader.Underline:
            case RtfTabLeader.ThickLine:
                return PdfCore.PdfTabLeaderStyle.Underscores;
            default:
                return PdfCore.PdfTabLeaderStyle.None;
        }
    }

    private static double ToNonNegativePoints(int? twips) =>
        twips.HasValue && twips.Value > 0 ? TwipsToPoints(twips.Value) : 0D;

    private static double ToSafeFirstLineIndent(int? twips, double leftIndentPoints) {
        if (!twips.HasValue) {
            return 0D;
        }

        double firstLineIndent = TwipsToPoints(twips.Value);
        return leftIndentPoints + firstLineIndent < 0D ? -leftIndentPoints : firstLineIndent;
    }

    private static double? ToPdfLineHeight(int? lineSpacingTwips, bool? multiple, double baseFontSize) {
        if (!lineSpacingTwips.HasValue || lineSpacingTwips.Value == 0) {
            return null;
        }

        double value = Math.Abs(lineSpacingTwips.Value);
        if (multiple == true) {
            return Math.Max(0.1D, value / 240D);
        }

        double lineHeightPoints = TwipsToPoints((int)value);
        return Math.Max(0.1D, lineHeightPoints / baseFontSize);
    }

    private static double GetParagraphBaseFontSize(RtfParagraph paragraph) {
        foreach (RtfRun run in paragraph.Runs) {
            if (run.FontSize.HasValue && run.FontSize.Value > 0) {
                return run.FontSize.Value;
            }
        }

        return 12D;
    }

    internal static double TwipsToPoints(int twips) => twips / 20D;
}
