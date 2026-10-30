using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static class RtfPdfMapping {
    private static readonly global::OfficeIMO.OfficeColor DefaultBorderColor = global::OfficeIMO.OfficeColor.FromRgb(0, 0, 0);

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

    internal static global::OfficeIMO.OfficeColor? ToPdfColor(RtfDocument document, int? oneBasedColorIndex) {
        RtfColor? color = oneBasedColorIndex.HasValue ? document.GetColor(oneBasedColorIndex.Value) : null;
        return color == null ? null : global::OfficeIMO.OfficeColor.FromRgb(color.Red, color.Green, color.Blue);
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

    internal static PdfCore.PdfParagraphStyle ToPdfParagraphStyle(RtfDocument document, RtfParagraph paragraph,
        IReadOnlyList<PdfCore.PdfTextRun> runs, PdfCore.PdfCreateOptions options) {
        paragraph = document.GetParagraphFormatting(paragraph);

        int? leftIndent = paragraph.LeftIndentTwips;
        int? rightIndent = paragraph.RightIndentTwips;
        int? firstLineIndent = paragraph.FirstLineIndentTwips;
        int? spaceBefore = paragraph.SpaceBeforeTwips;
        int? spaceAfter = paragraph.SpaceAfterTwips;
        bool? spaceBeforeAuto = paragraph.SpaceBeforeAuto;
        bool? spaceAfterAuto = paragraph.SpaceAfterAuto;
        int? lineSpacing = paragraph.LineSpacingTwips;
        bool? lineSpacingMultiple = paragraph.LineSpacingMultiple;
        bool keepTogether = paragraph.KeepLinesTogether;
        bool keepWithNext = paragraph.KeepWithNext;
        bool widowControl = paragraph.WidowControl ?? false;
        int? defaultTabWidth = document.Settings.DefaultTabWidthTwips;
        IReadOnlyList<RtfTabStop> tabStops = paragraph.TabStops;
        PdfCore.PdfParagraphStyle pdfStyle = new PdfCore.PdfParagraphStyle();
        double leftIndentPoints = ToNonNegativePoints(leftIndent);
        pdfStyle.LeftIndent = global::OfficeIMO.OfficeLength.Points(leftIndentPoints);
        pdfStyle.RightIndent = global::OfficeIMO.OfficeLength.Points(ToNonNegativePoints(rightIndent));
        pdfStyle.FirstLineIndent = global::OfficeIMO.OfficeLength.Points(ToSafeFirstLineIndent(firstLineIndent, leftIndentPoints));

        if (spaceBefore.HasValue && spaceBefore.Value >= 0) {
            pdfStyle.SpacingBefore = global::OfficeIMO.OfficeLength.Points(TwipsToPoints(spaceBefore.Value));
        } else if (spaceBeforeAuto == true) {
            pdfStyle.SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D);
        }

        if (spaceAfter.HasValue && spaceAfter.Value >= 0) {
            pdfStyle.SpacingAfter = global::OfficeIMO.OfficeLength.Points(TwipsToPoints(spaceAfter.Value));
        } else if (spaceAfterAuto == true) {
            pdfStyle.SpacingAfter = null;
        } else {
            // RTF's normal paragraph has no after gap. Do not inherit the PDF
            // builder's convenience gap when the source omits this control.
            pdfStyle.SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D);
        }

        pdfStyle.LineSpacing = ToPdfLineSpacing(lineSpacing, lineSpacingMultiple);
        bool exact = lineSpacing.HasValue && lineSpacing.Value < 0 && lineSpacingMultiple != true;
        if (!exact && PdfCore.PdfWriter.TryResolveImportedLineBoxFloor(runs,
                options.DefaultFont, options,
                out double layoutFontSize, out double naturalMultiplier)) {
            pdfStyle.FontSize = global::OfficeIMO.OfficeLength.Points(layoutFontSize);
            double multiple = lineSpacing.HasValue && lineSpacing.Value != 0 && lineSpacingMultiple == true
                ? Math.Max(0.1D, Math.Abs((double)lineSpacing.Value) / 240D) : 1D;
            pdfStyle.LineSpacing = lineSpacing.HasValue && lineSpacing.Value > 0 && lineSpacingMultiple != true
                ? PdfCore.PdfLineSpacing.AtLeast(global::OfficeIMO.OfficeLength.Points(lineSpacing.Value / 20D), naturalMultiplier)
                    .WithFontLineBoxBaseline(naturalMultiplier)
                : PdfCore.PdfLineSpacing.Multiple(naturalMultiplier * multiple)
                    .WithFontLineBoxBaseline(naturalMultiplier);
        }

        if (defaultTabWidth.HasValue && defaultTabWidth.Value > 0) {
            pdfStyle.DefaultTabStopWidth = global::OfficeIMO.OfficeLength.Points(TwipsToPoints(defaultTabWidth.Value));
        }

        foreach (RtfTabStop tabStop in tabStops) {
            pdfStyle.AddTabStop(global::OfficeIMO.OfficeLength.Points(TwipsToPoints(tabStop.PositionTwips)), ToPdfTabAlignment(tabStop.Alignment), ToPdfTabLeader(tabStop.Leader));
        }

        pdfStyle.KeepTogether = keepTogether;
        pdfStyle.KeepWithNext = keepWithNext;
        pdfStyle.WidowControl = widowControl;
        return pdfStyle;
    }

    internal static bool HasPageBreakBefore(RtfDocument document, RtfParagraph paragraph) {
        return document.GetParagraphFormatting(paragraph).PageBreakBefore;
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

        global::OfficeIMO.OfficeColor? color = ToPdfColor(document, source.ColorIndex);
        if (color.HasValue) {
            border.Color = color.Value;
        }

        return border;
    }

    internal static PdfCore.PdfTableStyle ToPdfTableStyle(RtfDocument document, RtfTable table, RtfToPdfOptions options) {
        PdfCore.PdfTableStyle? callerStyle = options.PdfOptions?.HasExplicitDefaultTableStyle == true
            ? options.PdfOptions.DefaultTableStyle
            : null;
        PdfCore.PdfTableStyle style = callerStyle?.Clone() ?? new PdfCore.PdfTableStyle {
            BorderColor = null,
            BorderWidth = global::OfficeIMO.OfficeLength.Points(0D),
            CellPaddingX = global::OfficeIMO.OfficeLength.Points(0D),
            CellPaddingY = global::OfficeIMO.OfficeLength.Points(0D)
        };
        int headerRowCount = GetRenderableHeaderRowCount(table);
        style.HeaderRowCount = headerRowCount;
        style.RepeatHeaderRowCount = headerRowCount;

        var fills = new Dictionary<(int Row, int Column), global::OfficeIMO.OfficeColor>();
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
                global::OfficeIMO.OfficeColor? fill = ToPdfColor(document, cell.BackgroundColorIndex ?? row.BackgroundColorIndex);
                if (fill.HasValue) {
                    fills[key] = fill.Value;
                }

                PdfCore.PdfCellPadding? padding = ToPdfCellPadding(row, cell, useRowGapFallback: callerStyle == null);
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

    private static PdfCore.PdfCellPadding? ToPdfCellPadding(RtfTableRow row, RtfTableCell cell, bool useRowGapFallback) {
        int? gap = useRowGapFallback && row.CellGapTwips.HasValue && row.CellGapTwips.Value >= 0
            ? row.CellGapTwips
            : null;
        double? top = ToOptionalNonNegativePoints(cell.PaddingTopTwips ?? row.PaddingTopTwips);
        double? left = ToOptionalNonNegativePoints(cell.PaddingLeftTwips ?? row.PaddingLeftTwips ?? gap);
        double? bottom = ToOptionalNonNegativePoints(cell.PaddingBottomTwips ?? row.PaddingBottomTwips);
        double? right = ToOptionalNonNegativePoints(cell.PaddingRightTwips ?? row.PaddingRightTwips ?? gap);
        if (!top.HasValue && !left.HasValue && !bottom.HasValue && !right.HasValue) {
            return null;
        }

        return new PdfCore.PdfCellPadding {
            Top = (top) is { } points12168 ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(points12168) : null,
            Left = (left) is { } points12192 ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(points12192) : null,
            Bottom = (bottom) is { } points12219 ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(points12219) : null,
            Right = (right) is { } points12247 ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(points12247) : null
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
            Width = global::OfficeIMO.OfficeLength.Points(width.HasValue && width.Value > 0 ? width.Value / 8D : 0.5D),
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

    private static PdfCore.PdfTabAlignment ToPdfTabAlignment(RtfTabAlignment alignment) {
        switch (alignment) {
            case RtfTabAlignment.Bar:
                return PdfCore.PdfTabAlignment.Bar;
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

    // RTF uses 240ths of a line in multiple mode. Otherwise the magnitude is
    // twips and its sign distinguishes a fixed advance from a minimum height.
    private static PdfCore.PdfLineSpacing? ToPdfLineSpacing(int? lineSpacingTwips, bool? multiple) {
        if (!lineSpacingTwips.HasValue || lineSpacingTwips.Value == 0) {
            return null;
        }

        double value = Math.Abs((double)lineSpacingTwips.Value);
        if (multiple == true) {
            return PdfCore.PdfLineSpacing.Multiple(Math.Max(0.1D, value / 240D));
        }

        double points = value / 20D;
        return lineSpacingTwips.Value < 0
            // Native exact-height RTF lines position the baseline within the
            // authored box independently of the selected run's font metrics.
            ? PdfCore.PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(points)).WithFixedLineBoxBaseline(points * 4D / 5D)
            : PdfCore.PdfLineSpacing.AtLeast(global::OfficeIMO.OfficeLength.Points(points));
    }

    internal static double TwipsToPoints(int twips) => twips / 20D;
}
