using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static class RtfPdfMapping {
    private static readonly PdfCore.PdfColor DefaultBorderColor = PdfCore.PdfColor.FromRgb(0, 0, 0);

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

    internal static PdfCore.PdfTableStyle ToPdfTableStyle(RtfDocument document, RtfTable table, RtfPdfSaveOptions options) {
        PdfCore.PdfTableStyle style = options.PdfOptions?.DefaultTableStyle?.Clone() ?? new PdfCore.PdfTableStyle();
        int headerRowCount = GetRenderableHeaderRowCount(table);
        style.HeaderRowCount = headerRowCount;
        style.RepeatHeaderRowCount = headerRowCount;

        var fills = new Dictionary<(int Row, int Column), PdfCore.PdfColor>();
        var paddings = new Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>();
        var verticalAlignments = new Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>();
        var borders = new Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>();

        int pdfRowIndex = 0;
        for (int sourceRowIndex = 0; sourceRowIndex < table.Rows.Count; sourceRowIndex++) {
            RtfTableRow row = table.Rows[sourceRowIndex];
            bool rowHasRenderableCells = false;
            for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++) {
                RtfTableCell cell = row.Cells[cellIndex];
                if (IsContinuationCell(cell)) {
                    continue;
                }

                rowHasRenderableCells = true;
                var key = (pdfRowIndex, cellIndex);
                PdfCore.PdfColor? fill = ToPdfColor(document, cell.BackgroundColorIndex ?? row.BackgroundColorIndex);
                if (fill.HasValue) {
                    fills[key] = fill.Value;
                }

                PdfCore.PdfCellPadding? padding = ToPdfCellPadding(row, cell);
                if (padding != null) {
                    paddings[key] = padding;
                }

                PdfCore.PdfCellVerticalAlign? verticalAlignment = ToPdfCellVerticalAlign(cell.VerticalAlignment);
                if (verticalAlignment.HasValue) {
                    verticalAlignments[key] = verticalAlignment.Value;
                }

                PdfCore.PdfCellBorder? border = ToPdfCellBorder(document, row, cell);
                if (border != null) {
                    borders[key] = border;
                }
            }

            if (rowHasRenderableCells) {
                pdfRowIndex++;
            }
        }

        if (fills.Count > 0) style.CellFills = fills;
        if (paddings.Count > 0) style.CellPaddings = paddings;
        if (verticalAlignments.Count > 0) style.CellVerticalAlignments = verticalAlignments;
        if (borders.Count > 0) style.CellBorders = borders;
        return style;
    }

    private static RtfPageBorder? GetFirstRenderablePageBorder(RtfPageBorders borders) {
        if (IsRenderablePageBorder(borders.Top)) return borders.Top;
        if (IsRenderablePageBorder(borders.Bottom)) return borders.Bottom;
        if (IsRenderablePageBorder(borders.Left)) return borders.Left;
        if (IsRenderablePageBorder(borders.Right)) return borders.Right;
        return null;
    }

    private static bool IsRenderablePageBorder(RtfPageBorder border) => border.Style != RtfPageBorderStyle.None;

    private static int GetRenderableHeaderRowCount(RtfTable table) {
        int headerRowCount = 0;
        foreach (RtfTableRow row in table.Rows) {
            if (!row.RepeatHeader) {
                break;
            }

            if (HasRenderableCells(row)) {
                headerRowCount++;
            }
        }

        return headerRowCount;
    }

    private static bool HasRenderableCells(RtfTableRow row) =>
        row.Cells.Any(cell => !IsContinuationCell(cell));

    private static bool IsContinuationCell(RtfTableCell cell) =>
        cell.HorizontalMerge == RtfTableCellMerge.Continue ||
        cell.VerticalMerge == RtfTableCellMerge.Continue;

    private static PdfCore.PdfCellPadding? ToPdfCellPadding(RtfTableRow row, RtfTableCell cell) {
        double? top = ToOptionalNonNegativePoints(cell.PaddingTopTwips ?? row.PaddingTopTwips);
        double? left = ToOptionalNonNegativePoints(cell.PaddingLeftTwips ?? row.PaddingLeftTwips);
        double? bottom = ToOptionalNonNegativePoints(cell.PaddingBottomTwips ?? row.PaddingBottomTwips);
        double? right = ToOptionalNonNegativePoints(cell.PaddingRightTwips ?? row.PaddingRightTwips);
        if (!top.HasValue && !left.HasValue && !bottom.HasValue && !right.HasValue) {
            return null;
        }

        return new PdfCore.PdfCellPadding {
            Top = top,
            Left = left,
            Bottom = bottom,
            Right = right
        };
    }

    private static PdfCore.PdfCellVerticalAlign? ToPdfCellVerticalAlign(RtfTableCellVerticalAlignment? alignment) {
        switch (alignment) {
            case RtfTableCellVerticalAlignment.Top:
                return PdfCore.PdfCellVerticalAlign.Top;
            case RtfTableCellVerticalAlignment.Center:
                return PdfCore.PdfCellVerticalAlign.Middle;
            case RtfTableCellVerticalAlignment.Bottom:
                return PdfCore.PdfCellVerticalAlign.Bottom;
            default:
                return null;
        }
    }

    private static PdfCore.PdfCellBorder? ToPdfCellBorder(RtfDocument document, RtfTableRow row, RtfTableCell cell) {
        PdfCore.PdfCellBorderSide? top = ToPdfCellBorderSide(document, cell.TopBorder, row.TopBorder);
        PdfCore.PdfCellBorderSide? right = ToPdfCellBorderSide(document, cell.RightBorder, row.RightBorder);
        PdfCore.PdfCellBorderSide? bottom = ToPdfCellBorderSide(document, cell.BottomBorder, row.BottomBorder);
        PdfCore.PdfCellBorderSide? left = ToPdfCellBorderSide(document, cell.LeftBorder, row.LeftBorder);
        PdfCore.PdfCellBorderSide? diagonalDown = ToPdfDiagonalCellBorderSide(document, cell.TopLeftToBottomRightBorder);
        PdfCore.PdfCellBorderSide? diagonalUp = ToPdfDiagonalCellBorderSide(document, cell.TopRightToBottomLeftBorder);
        if (top == null && right == null && bottom == null && left == null && diagonalDown == null && diagonalUp == null) {
            return null;
        }

        return new PdfCore.PdfCellBorder {
            Top = top != null,
            Right = right != null,
            Bottom = bottom != null,
            Left = left != null,
            DiagonalDown = diagonalDown != null,
            DiagonalUp = diagonalUp != null,
            TopBorder = top,
            RightBorder = right,
            BottomBorder = bottom,
            LeftBorder = left,
            DiagonalDownBorder = diagonalDown,
            DiagonalUpBorder = diagonalUp,
            Color = DefaultBorderColor
        };
    }

    private static PdfCore.PdfCellBorderSide? ToPdfCellBorderSide(RtfDocument document, RtfTableCellBorder cellBorder, RtfTableRowBorder? rowBorder) {
        RtfTableCellBorderStyle style = cellBorder.HasAnyValue ? cellBorder.Style : rowBorder?.Style ?? RtfTableCellBorderStyle.None;
        if (style == RtfTableCellBorderStyle.None) {
            return null;
        }

        int? width = cellBorder.HasAnyValue ? cellBorder.Width : rowBorder?.Width;
        int? colorIndex = cellBorder.HasAnyValue ? cellBorder.ColorIndex : rowBorder?.ColorIndex;
        return CreatePdfCellBorderSide(document, style, width, colorIndex);
    }

    private static PdfCore.PdfCellBorderSide? ToPdfDiagonalCellBorderSide(RtfDocument document, RtfTableCellBorder cellBorder) {
        if (!cellBorder.HasAnyValue || cellBorder.Style == RtfTableCellBorderStyle.None) {
            return null;
        }

        return CreatePdfCellBorderSide(document, cellBorder.Style, cellBorder.Width, cellBorder.ColorIndex);
    }

    private static PdfCore.PdfCellBorderSide CreatePdfCellBorderSide(RtfDocument document, RtfTableCellBorderStyle style, int? width, int? colorIndex) {
        return new PdfCore.PdfCellBorderSide {
            Color = ToPdfColor(document, colorIndex) ?? DefaultBorderColor,
            Width = width.HasValue && width.Value > 0 ? width.Value / 8D : 0.5D,
            DashStyle = ToPdfDashStyle(style),
            LineStyle = style == RtfTableCellBorderStyle.Double ? PdfCore.PdfCellBorderLineStyle.TwoLine : PdfCore.PdfCellBorderLineStyle.Standard
        };
    }

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

    private static OfficeIMO.Drawing.OfficeStrokeDashStyle ToPdfDashStyle(RtfTableCellBorderStyle style) {
        switch (style) {
            case RtfTableCellBorderStyle.Dashed:
                return OfficeIMO.Drawing.OfficeStrokeDashStyle.Dash;
            case RtfTableCellBorderStyle.Dotted:
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

    private static double? ToOptionalNonNegativePoints(int? twips) =>
        twips.HasValue && twips.Value >= 0 ? TwipsToPoints(twips.Value) : (double?)null;

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
