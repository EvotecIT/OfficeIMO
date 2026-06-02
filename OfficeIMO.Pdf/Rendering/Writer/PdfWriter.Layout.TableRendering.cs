using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static bool CanRenderTableCellCheckBoxInline(TableCellLayout cell, TableCellTextLayout layout, int startLine, int lineCount) =>
        startLine == 0 &&
        lineCount > 0 &&
        cell.Images.Count == 0 &&
        cell.FormFields.Count == 0 &&
        cell.CheckBoxes.Count == 1 &&
        !string.IsNullOrWhiteSpace(cell.Text) &&
        layout.Lines.Count == 1 &&
        layout.Lines[0].Count > 0;

    private static void RenderTableCellInlineCheckBox(LayoutResult.Page page, TableCellLayout cell, PdfColumnAlign align, System.Collections.Generic.IReadOnlyList<RichSeg> line, double textX, double innerWidth, double baselineY) {
        PdfTableCellCheckBox checkBox = cell.CheckBoxes[0];
        double size = checkBox.Size;
        double lineWidth = MeasureRichLineWidth(line);
        double lineX = align switch {
            PdfColumnAlign.Center => textX + System.Math.Max(0D, (innerWidth - lineWidth) / 2D),
            PdfColumnAlign.Right => textX + System.Math.Max(0D, innerWidth - lineWidth),
            _ => textX
        };
        double x = System.Math.Min(textX + System.Math.Max(0D, innerWidth - size), lineX + lineWidth + TableCellCheckBoxGap);
        double topY = baselineY + (size * 0.75D);
        page.FormFields.Add(new FormFieldAnnotation {
            X1 = x,
            Y1 = topY - size,
            X2 = x + size,
            Y2 = topY,
            Kind = FormFieldAnnotationKind.CheckBox,
            Name = checkBox.Name,
            Value = checkBox.IsChecked ? checkBox.CheckedValueName : "Off",
            IsChecked = checkBox.IsChecked,
            CheckedValueName = checkBox.CheckedValueName,
            Style = checkBox.Style
        });
    }

    private static void RenderTableCellObjects(LayoutResult.Page page, TableCellLayout cell, PdfColumnAlign align, double textX, double innerWidth, double topY) {
        double yCursor = topY;
        int objectCount = 0;
        for (int index = 0; index < cell.Images.Count; index++) {
            PdfTableCellImage image = cell.Images[index];
            if (objectCount > 0) {
                yCursor -= TableCellCheckBoxGap;
            }

            PdfAlign imageAlign = image.Style?.Align ?? MapTableCellAlignment(align);
            ImageBlock block = image.ToImageBlock(imageAlign);
            PdfImageStyle imageStyle = block.Style ?? new PdfImageStyle {
                Align = imageAlign
            };
            PdfDoc.ValidateImageStyleForBox(imageStyle, block.Width, block.Height, nameof(image.Style));
            PdfDoc.ValidateImageFitDimensions(block.Info, imageStyle.Fit, nameof(image.Style));
            double x = imageStyle.Align switch {
                PdfAlign.Center => textX + System.Math.Max(0D, (innerWidth - block.Width) / 2D),
                PdfAlign.Right => textX + System.Math.Max(0D, innerWidth - block.Width),
                _ => textX
            };
            PageImage pageImage = CreatePageImage(block, imageStyle, x, yCursor - block.Height);
            page.Images.Add(pageImage);
            AddTableCellImageLinkAnnotation(page, image, imageStyle, pageImage, x, yCursor - block.Height);
            yCursor -= block.Height;
            objectCount++;
        }

        for (int index = 0; index < cell.CheckBoxes.Count; index++) {
            PdfTableCellCheckBox checkBox = cell.CheckBoxes[index];
            if (objectCount > 0) {
                yCursor -= TableCellCheckBoxGap;
            }

            double size = checkBox.Size;
            double x = align switch {
                PdfColumnAlign.Center => textX + (innerWidth - size) / 2D,
                PdfColumnAlign.Right => textX + innerWidth - size,
                _ => textX
            };
            page.FormFields.Add(new FormFieldAnnotation {
                X1 = x,
                Y1 = yCursor - size,
                X2 = x + size,
                Y2 = yCursor,
                Kind = FormFieldAnnotationKind.CheckBox,
                Name = checkBox.Name,
                Value = checkBox.IsChecked ? checkBox.CheckedValueName : "Off",
                IsChecked = checkBox.IsChecked,
                CheckedValueName = checkBox.CheckedValueName,
                Style = checkBox.Style
            });
            yCursor -= size;
            objectCount++;
        }

        for (int index = 0; index < cell.FormFields.Count; index++) {
            PdfTableCellFormField formField = cell.FormFields[index];
            if (objectCount > 0) {
                yCursor -= TableCellCheckBoxGap;
            }

            double width = System.Math.Min(formField.Width, innerWidth);
            double x = align switch {
                PdfColumnAlign.Center => textX + (innerWidth - width) / 2D,
                PdfColumnAlign.Right => textX + innerWidth - width,
                _ => textX
            };

            page.FormFields.Add(new FormFieldAnnotation {
                X1 = x,
                Y1 = yCursor - formField.Height,
                X2 = x + width,
                Y2 = yCursor,
                Kind = formField.Kind == PdfTableCellFormFieldKind.Text ? FormFieldAnnotationKind.Text : FormFieldAnnotationKind.Choice,
                Name = formField.Name,
                Value = formField.Value,
                Values = formField.Values,
                FontSize = formField.FontSize,
                Options = formField.Options,
                IsComboBox = formField.IsComboBox,
                AllowsMultipleSelection = false,
                Style = formField.Style
            });
            yCursor -= formField.Height;
            objectCount++;
        }
    }

    private static void AddTableCellImageLinkAnnotation(LayoutResult.Page page, PdfTableCellImage image, PdfImageStyle style, PageImage pageImage, double targetX, double targetBottomY) {
        if (string.IsNullOrEmpty(image.LinkUri)) {
            return;
        }

        double x1 = pageImage.X;
        double y1 = pageImage.Y;
        double x2 = pageImage.X + pageImage.W;
        double y2 = pageImage.Y + pageImage.H;
        if (style.Fit == OfficeImageFit.Cover || style.ClipPath != null) {
            x1 = targetX;
            y1 = targetBottomY;
            x2 = targetX + image.Width;
            y2 = targetBottomY + image.Height;
        }

        page.Annotations.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = image.LinkUri!, Contents = image.LinkContents });
    }

    private static System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> SliceTableCellLines(TableCellTextLayout layout, int startLine, int lineCount) {
        var lines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>();
        int available = System.Math.Max(0, layout.Lines.Count - startLine);
        int visible = System.Math.Max(0, System.Math.Min(lineCount, available));
        for (int i = 0; i < visible; i++) {
            lines.Add(layout.Lines[startLine + i]);
        }

        if (lines.Count == 0) {
            lines.Add(new System.Collections.Generic.List<RichSeg>());
        }

        return lines;
    }

    private static System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> StripRichLineLinksWhenCellLinked(System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> lines, string? linkUri, string? linkDestinationName) {
        if (!HasCellLinkTarget(linkUri, linkDestinationName) || !lines.Any(line => line.Any(segment => segment.Uri != null || segment.DestinationName != null))) {
            return lines;
        }

        var stripped = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>(lines.Count);
        foreach (System.Collections.Generic.List<RichSeg> line in lines) {
            var strippedLine = new System.Collections.Generic.List<RichSeg>(line.Count);
            foreach (RichSeg segment in line) {
                strippedLine.Add(segment.WithoutLink());
            }

            stripped.Add(strippedLine);
        }

        return stripped;
    }

    private static System.Collections.Generic.List<double> SliceTableCellLineHeights(TableCellTextLayout layout, int startLine, int lineCount, double fallbackLeading) {
        var heights = new System.Collections.Generic.List<double>();
        int available = System.Math.Max(0, layout.Lines.Count - startLine);
        int visible = System.Math.Max(0, System.Math.Min(lineCount, available));
        for (int i = 0; i < visible; i++) {
            int lineIndex = startLine + i;
            heights.Add(lineIndex < layout.LineHeights.Count ? layout.LineHeights[lineIndex] : fallbackLeading);
        }

        if (heights.Count == 0) {
            heights.Add(fallbackLeading);
        }

        return heights;
    }

    private static PdfAlign MapTableCellAlignment(PdfColumnAlign align) => align switch {
        PdfColumnAlign.Center => PdfAlign.Center,
        PdfColumnAlign.Right => PdfAlign.Right,
        _ => PdfAlign.Left
    };

    private static bool[] GetRowSpanBoundarySkipColumns(TableBlock table, int boundaryRowIndex, int columnCount) {
        var skipColumns = new bool[columnCount];
        if (boundaryRowIndex < 0 || boundaryRowIndex >= table.Cells.Count - 1 || columnCount <= 0) {
            return skipColumns;
        }

        for (int rowIndex = 0; rowIndex <= boundaryRowIndex; rowIndex++) {
            var cells = GetTableCellLayouts(table, rowIndex, columnCount);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCellLayout cell = cells[cellIndex];
                if (cell.RowSpan <= 1 || boundaryRowIndex >= rowIndex + cell.RowSpan - 1) {
                    continue;
                }

                int lastColumn = System.Math.Min(columnCount, cell.Column + cell.ColumnSpan);
                for (int column = cell.Column; column < lastColumn; column++) {
                    skipColumns[column] = true;
                }
            }
        }

        return skipColumns;
    }

    private static bool[] GetRowSpanContinuationSkipColumns(TableBlock table, int rowIndex, int columnCount) {
        var skipColumns = new bool[columnCount];
        if (rowIndex <= 0 || rowIndex >= table.Cells.Count || columnCount <= 0) {
            return skipColumns;
        }

        for (int startRow = 0; startRow < rowIndex; startRow++) {
            var cells = GetTableCellLayouts(table, startRow, columnCount);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCellLayout cell = cells[cellIndex];
                if (cell.RowSpan <= rowIndex - startRow) {
                    continue;
                }

                int lastColumn = System.Math.Min(columnCount, cell.Column + cell.ColumnSpan);
                for (int column = cell.Column; column < lastColumn; column++) {
                    skipColumns[column] = true;
                }
            }
        }

        return skipColumns;
    }

    private static bool[] GetMergedCellContinuationSkipColumns(TableBlock table, int rowIndex, int columnCount) {
        bool[] skipColumns = GetRowSpanContinuationSkipColumns(table, rowIndex, columnCount);
        if (rowIndex < 0 || rowIndex >= table.Cells.Count || columnCount <= 0) {
            return skipColumns;
        }

        var cells = GetTableCellLayouts(table, rowIndex, columnCount);
        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
            TableCellLayout cell = cells[cellIndex];
            int lastColumn = System.Math.Min(columnCount, cell.Column + cell.ColumnSpan);
            for (int column = cell.Column + 1; column < lastColumn; column++) {
                skipColumns[column] = true;
            }
        }

        return skipColumns;
    }

    private static void DrawTableHorizontalLine(StringBuilder sb, PdfColor color, double width, double xOrigin, double[] columnWidths, double columnGap, double y, bool[] skipColumns) {
        if (columnWidths.Length == 0) {
            return;
        }

        double tableWidth = GetTableCellWidth(columnWidths, 0, columnWidths.Length, columnGap);
        if (!HasSkippedColumns(skipColumns, columnWidths.Length)) {
            DrawHLine(sb, color, width, xOrigin, xOrigin + tableWidth, y);
            return;
        }

        var columnLefts = new double[columnWidths.Length];
        var columnRights = new double[columnWidths.Length];
        double x = xOrigin;
        for (int column = 0; column < columnWidths.Length; column++) {
            columnLefts[column] = x;
            columnRights[column] = x + columnWidths[column];
            x += columnWidths[column] + columnGap;
        }

        int segmentStart = -1;
        for (int column = 0; column < columnWidths.Length; column++) {
            bool skip = column < skipColumns.Length && skipColumns[column];
            if (skip) {
                if (segmentStart >= 0) {
                    DrawHLine(sb, color, width, columnLefts[segmentStart], columnRights[column - 1], y);
                    segmentStart = -1;
                }

                continue;
            }

            if (segmentStart < 0) {
                segmentStart = column;
            }
        }

        if (segmentStart >= 0) {
            DrawHLine(sb, color, width, columnLefts[segmentStart], columnRights[columnWidths.Length - 1], y);
        }
    }

    private static void DrawTableRowFill(StringBuilder sb, PdfColor color, double xOrigin, double[] columnWidths, double columnGap, double y, double height, bool[] skipColumns) {
        if (columnWidths.Length == 0) {
            return;
        }

        double tableWidth = GetTableCellWidth(columnWidths, 0, columnWidths.Length, columnGap);
        if (!HasSkippedColumns(skipColumns, columnWidths.Length)) {
            DrawRowFill(sb, color, xOrigin, y, tableWidth, height);
            return;
        }

        var columnLefts = new double[columnWidths.Length];
        var columnRights = new double[columnWidths.Length];
        double x = xOrigin;
        for (int column = 0; column < columnWidths.Length; column++) {
            columnLefts[column] = x;
            columnRights[column] = x + columnWidths[column];
            x += columnWidths[column] + columnGap;
        }

        int segmentStart = -1;
        for (int column = 0; column < columnWidths.Length; column++) {
            bool skip = column < skipColumns.Length && skipColumns[column];
            if (skip) {
                if (segmentStart >= 0) {
                    DrawRowFill(sb, color, columnLefts[segmentStart], y, columnRights[column - 1] - columnLefts[segmentStart], height);
                    segmentStart = -1;
                }

                continue;
            }

            if (segmentStart < 0) {
                segmentStart = column;
            }
        }

        if (segmentStart >= 0) {
            DrawRowFill(sb, color, columnLefts[segmentStart], y, columnRights[columnWidths.Length - 1] - columnLefts[segmentStart], height);
        }
    }

    private static bool DrawTableCellDataBars(StringBuilder sb, PdfTableStyle style, System.Collections.Generic.List<TableCellLayout> cells, int rowIndex, int columnCount, double xOrigin, double yTop, double rowBottom, double rowHeight, double[] columnWidths, double columnGap, double[] rowHeights, double rowGap, bool wholeRowSegment, int startLine, bool[] skipColumns) {
        if (style.CellDataBars == null || style.CellDataBars.Count == 0 || startLine != 0) {
            return false;
        }

        bool drawn = false;
        double cellX = xOrigin;
        for (int column = 0; column < columnCount; column++) {
            if (style.CellDataBars.TryGetValue((rowIndex, column), out PdfCellDataBar? dataBar) &&
                dataBar.Ratio > 0D &&
                TryGetTableCellLayoutAtColumn(cells, column, out TableCellLayout cell) &&
                (column >= skipColumns.Length || !skipColumns[column])) {
                int span = wholeRowSegment ? cell.ColumnSpan : 1;
                double cellWidth = GetTableCellWidth(columnWidths, column, span, columnGap);
                double cellHeight = rowHeight;
                double cellBottom = rowBottom;
                if (wholeRowSegment && cell.RowSpan > 1) {
                    cellHeight = GetTableCellHeight(rowHeights, rowIndex, cell.RowSpan, rowGap);
                    cellBottom = yTop - cellHeight;
                }

                double padLeft = GetTableCellPaddingLeft(style, rowIndex, column);
                double padRight = GetTableCellPaddingRight(style, rowIndex, column);
                double padTop = GetTableCellPaddingTop(style, rowIndex, column);
                double padBottom = GetTableCellPaddingBottom(style, rowIndex, column);
                double contentWidth = System.Math.Max(0D, cellWidth - padLeft - padRight);
                double barX = cellX + padLeft + contentWidth * dataBar.StartRatio;
                double barWidth = contentWidth * dataBar.Ratio;
                double barHeight = System.Math.Max(0D, cellHeight - padTop - padBottom);
                if (barWidth > 0.001D && barHeight > 0.001D) {
                    DrawRowFill(sb, dataBar.Color, barX, cellBottom + padBottom, barWidth, barHeight);
                    drawn = true;
                }
            }

            cellX += columnWidths[column] + columnGap;
        }

        return drawn;
    }

    private static bool DrawTableCellIcons(StringBuilder sb, PdfTableStyle style, System.Collections.Generic.List<TableCellLayout> cells, int rowIndex, int columnCount, double xOrigin, double yTop, double rowBottom, double rowHeight, double[] columnWidths, double columnGap, double[] rowHeights, double rowGap, bool wholeRowSegment, int startLine, bool[] skipColumns) {
        if (style.CellIcons == null || style.CellIcons.Count == 0 || startLine != 0) {
            return false;
        }

        bool drawn = false;
        double cellX = xOrigin;
        for (int column = 0; column < columnCount; column++) {
            if (style.CellIcons.TryGetValue((rowIndex, column), out PdfCellIcon? icon) &&
                TryGetTableCellLayoutAtColumn(cells, column, out TableCellLayout cell) &&
                (column >= skipColumns.Length || !skipColumns[column])) {
                int span = wholeRowSegment ? cell.ColumnSpan : 1;
                double cellWidth = GetTableCellWidth(columnWidths, column, span, columnGap);
                double cellHeight = rowHeight;
                double cellBottom = rowBottom;
                if (wholeRowSegment && cell.RowSpan > 1) {
                    cellHeight = GetTableCellHeight(rowHeights, rowIndex, cell.RowSpan, rowGap);
                    cellBottom = yTop - cellHeight;
                }

                double iconSize = Math.Min(icon.Size, Math.Max(1D, Math.Min(cellWidth, cellHeight) - 2D));
                if (iconSize > 0.001D) {
                    double padLeft = GetTableCellPaddingLeft(style, rowIndex, column);
                    double padRight = GetTableCellPaddingRight(style, rowIndex, column);
                    double padTop = GetTableCellPaddingTop(style, rowIndex, column);
                    double padBottom = GetTableCellPaddingBottom(style, rowIndex, column);
                    double contentLeft = cellX + padLeft;
                    double contentRight = cellX + cellWidth - padRight;
                    double contentBottom = cellBottom + padBottom;
                    double contentTop = cellBottom + cellHeight - padTop;
                    double contentWidth = Math.Max(0D, contentRight - contentLeft);
                    double contentHeight = Math.Max(0D, contentTop - contentBottom);
                    PdfColumnAlign horizontalAlign = GetTableCellAlignment(style, rowIndex, column, cell.Text);
                    PdfCellVerticalAlign verticalAlign = GetTableCellVerticalAlignment(style, rowIndex, column);
                    double iconX = horizontalAlign switch {
                        PdfColumnAlign.Center => contentLeft + Math.Max(0D, (contentWidth - iconSize) / 2D),
                        PdfColumnAlign.Right => contentRight - iconSize,
                        _ => contentLeft
                    };
                    double iconY = verticalAlign switch {
                        PdfCellVerticalAlign.Middle => contentBottom + Math.Max(0D, (contentHeight - iconSize) / 2D),
                        PdfCellVerticalAlign.Bottom => contentBottom,
                        _ => contentTop - iconSize
                    };

                    iconX += icon.OffsetX;
                    iconY += icon.OffsetY;
                    DrawTableCellIcon(sb, icon, iconX, iconY, iconSize);
                    drawn = true;
                }
            }

            cellX += columnWidths[column] + columnGap;
        }

        return drawn;
    }

    private static void DrawTableCellIcon(StringBuilder sb, PdfCellIcon icon, double x, double y, double size) {
        var content = new ContentStreamBuilder(sb);
        content.FillColor(icon.Color);
        double midX = x + size / 2D;
        double midY = y + size / 2D;
        switch (icon.Kind) {
            case PdfCellIconKind.Circle:
                DrawFilledCircle(content, midX, midY, size / 2D);
                break;
            case PdfCellIconKind.Diamond:
                content.MoveTo(midX, y + size).LineTo(x + size, midY).LineTo(midX, y).LineTo(x, midY).ClosePath().FillPath();
                break;
            case PdfCellIconKind.Square:
                content.Rectangle(x, y, size, size).FillPath();
                break;
            case PdfCellIconKind.TriangleUp:
                content.MoveTo(midX, y + size).LineTo(x + size, y).LineTo(x, y).ClosePath().FillPath();
                break;
            case PdfCellIconKind.TriangleRight:
                content.MoveTo(x + size, midY).LineTo(x, y + size).LineTo(x, y).ClosePath().FillPath();
                break;
            case PdfCellIconKind.TriangleDown:
                content.MoveTo(x, y + size).LineTo(x + size, y + size).LineTo(midX, y).ClosePath().FillPath();
                break;
            case PdfCellIconKind.CheckBoxUnchecked:
                DrawCheckBoxIcon(content, icon.Color, x, y, size, selected: false);
                break;
            case PdfCellIconKind.CheckBoxChecked:
                DrawCheckBoxIcon(content, icon.Color, x, y, size, selected: true);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(icon), icon.Kind, "PDF table cell icon kind is not supported.");
        }
    }

    private static void DrawCheckBoxIcon(ContentStreamBuilder content, PdfColor color, double x, double y, double size, bool selected) {
        double strokeWidth = Math.Max(0.75D, size * 0.085D);
        double inset = strokeWidth / 2D;
        content.StrokeColor(color)
            .LineWidth(strokeWidth)
            .Rectangle(x + inset, y + inset, size - (inset * 2D), size - (inset * 2D))
            .StrokePath();

        if (!selected) {
            return;
        }

        content.StrokeColor(color)
            .LineWidth(Math.Max(1D, size * 0.14D))
            .LineCap(1)
            .LineJoin(1)
            .MoveTo(x + (size * 0.24D), y + (size * 0.52D))
            .LineTo(x + (size * 0.43D), y + (size * 0.31D))
            .LineTo(x + (size * 0.78D), y + (size * 0.72D))
            .StrokePath()
            .LineCap(0)
            .LineJoin(0);
    }

    private static void DrawFilledCircle(ContentStreamBuilder content, double centerX, double centerY, double radius) {
        const double kappa = 0.5522847498307936D;
        double control = radius * kappa;
        content.MoveTo(centerX + radius, centerY)
            .CubicTo(centerX + radius, centerY + control, centerX + control, centerY + radius, centerX, centerY + radius)
            .CubicTo(centerX - control, centerY + radius, centerX - radius, centerY + control, centerX - radius, centerY)
            .CubicTo(centerX - radius, centerY - control, centerX - control, centerY - radius, centerX, centerY - radius)
            .CubicTo(centerX + control, centerY - radius, centerX + radius, centerY - control, centerX + radius, centerY)
            .ClosePath()
            .FillPath();
    }

    private static bool HasSkippedColumns(bool[] skipColumns, int columnCount) {
        for (int column = 0; column < columnCount && column < skipColumns.Length; column++) {
            if (skipColumns[column]) {
                return true;
            }
        }

        return false;
    }

    private static bool ShouldClipTableCellText(double textX, double textBaselineY, double textWidth, PdfStandardFont font, double fontSize, double cellX, double cellY, double cellWidth, double cellHeight) {
        const double epsilon = 0.01D;
        double ascender = GetAscender(font, fontSize);
        double descender = GetDescender(font, fontSize);

        return textX < cellX - epsilon ||
               textX + textWidth > cellX + cellWidth + epsilon ||
               textBaselineY + ascender > cellY + cellHeight + epsilon ||
               textBaselineY + descender < cellY - epsilon;
    }

    private static double GetTableRowFontSize(PdfTableStyle style, int rowIndex, int headerRowCount, int footerStartRowIndex, double defaultFontSize) {
        if (rowIndex < headerRowCount) {
            return style.HeaderFontSize ?? GetTableBodyFontSize(style, defaultFontSize);
        }

        if (rowIndex >= footerStartRowIndex) {
            return style.FooterFontSize ?? GetTableBodyFontSize(style, defaultFontSize);
        }

        return GetTableBodyFontSize(style, defaultFontSize);
    }

    private static bool GetTableRowBold(PdfTableStyle style, int rowIndex, int headerRowCount, int footerStartRowIndex) {
        return rowIndex < headerRowCount ? style.HeaderBold : rowIndex >= footerStartRowIndex && style.FooterBold;
    }

    private static PdfStandardFont GetTableRowFont(PdfOptions options, bool bold) {
        var normalFont = ChooseNormal(options.DefaultFont);
        return bold ? ChooseBold(normalFont) : normalFont;
    }

    private static string GetTableRowFontResource(bool bold) {
        return bold ? "F2" : "F1";
    }

    private static bool TableUsesBold(PdfTableStyle style, int rowCount, int headerRowCount, int footerStartRowIndex) {
        for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
            if (GetTableRowBold(style, rowIndex, headerRowCount, footerStartRowIndex)) {
                return true;
            }
        }

        return false;
    }

}
