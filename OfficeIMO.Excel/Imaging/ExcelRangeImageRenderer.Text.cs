using System;
using System.Collections.Generic;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    internal static partial class ExcelRangeImageRenderer {
        private const double CellTextFontSize = 11D;
        private const double CellTextHorizontalPadding = 3D;
        private const double CellTextVerticalPadding = 1D;
        private const double CellTextLineHeightFactor = 1.2D;
        private static readonly OfficeColor HyperlinkHintColor = OfficeColor.FromRgb(5, 99, 193);

        private static void DrawRasterCellText(
            OfficeRasterCanvas canvas,
            ExcelVisualCell cell,
            ExcelRangeVisualSnapshot snapshot,
            ExcelImageExportOptions options,
            double scale,
            List<OfficeImageExportDiagnostic>? diagnostics) {
            if (string.IsNullOrEmpty(cell.Text)) {
                return;
            }

            double x = cell.X * scale;
            double y = cell.Y * scale;
            double w = cell.Width * scale;
            double h = cell.Height * scale;
            double paddingX = CellTextHorizontalPadding * scale;
            double paddingY = CellTextVerticalPadding * scale;
            double availableWidth = Math.Max(1D, w - (paddingX * 2D));
            double availableHeight = Math.Max(1D, h - (paddingY * 2D));
            double fontSize = ResolveCellFontSize(cell.Style, scale);
            double minimumFontSize = Math.Max(1D, scale);
            double rotationDegrees = ResolveExcelTextRotationDegrees(cell.Style.TextRotation, snapshot, cell, diagnostics);
            bool rotated = Math.Abs(rotationDegrees) > 0.0001D;
            bool richTextSupported = IsRichTextRenderingSupported(cell, rotated);
            OfficeTextBlockLayout layout = OfficeTextLayoutEngine.LayoutTextBlock(
                cell.Text,
                fontSize,
                rotated ? Math.Max(availableWidth, availableHeight) : availableWidth,
                rotated ? Math.Max(availableWidth, availableHeight) : availableHeight,
                CellTextLineHeightFactor,
                minimumFontSize,
                canvas.MeasureText,
                wrap: cell.Style.WrapText,
                forceSingleLine: rotated,
                shrinkToFit: cell.Style.ShrinkToFit);
            if (layout.Lines.Count == 0) {
                return;
            }

            using (canvas.PushClipRectangle(x, y, w, h)) {
                if (cell.RichTextRuns.Count > 0) {
                    if (richTextSupported && TryDrawRasterRichText(canvas, cell, options, scale, x, y, w, h, paddingX, paddingY, availableWidth, availableHeight, rotationDegrees, out OfficeRichTextBlockLayout richLayout)) {
                        AddTextClippingDiagnosticIfNeeded(richLayout, snapshot, cell, diagnostics);
                        if (rotated) {
                            AddRotatedTextApproximationDiagnostic(snapshot, cell, diagnostics);
                        }

                        return;
                    }

                    AddRichTextLayoutApproximationDiagnostic(snapshot, cell, diagnostics);
                }

                AddTextClippingDiagnosticIfNeeded(layout, snapshot, cell, diagnostics);
                OfficeColor color = ResolveCellTextColor(cell, options);
                OfficeTextAlignment alignment = ResolveTextAlignment(cell.Style.HorizontalAlignment);
                bool bold = cell.Style.Bold;
                bool italic = cell.Style.Italic;
                bool underline = ShouldUnderlineText(cell, options);
                if (rotated) {
                    AddRotatedTextApproximationDiagnostic(snapshot, cell, diagnostics);
                    OfficeTextLine line = layout.Lines[0];
                    double centerX = x + (w / 2D);
                    double centerY = y + (h / 2D);
                    double textTop = centerY - (layout.FontSize / 2D);
                    canvas.DrawTextLine(line.Text, centerX, textTop, layout.FontSize, color, bold, italic, OfficeTextAlignment.Center, rotationDegrees, centerX, centerY);
                    return;
                }

                OfficeTextBlockRenderer.DrawRasterTextBlock(
                    canvas,
                    layout,
                    x + paddingX,
                    y + paddingY,
                    availableWidth,
                    availableHeight,
                    color,
                    alignment,
                    ResolveTextVerticalAlignment(cell.Style.VerticalAlignment),
                    bold,
                    italic,
                    underline);
            }
        }

        private static void AppendSvgCellText(
            StringBuilder builder,
            ExcelVisualCell cell,
            ExcelRangeVisualSnapshot snapshot,
            ExcelImageExportOptions options,
            OfficeRasterCanvas textMeasureCanvas,
            List<OfficeImageExportDiagnostic>? diagnostics) {
            if (string.IsNullOrEmpty(cell.Text)) {
                return;
            }

            double scale = options.Scale;
            double x = cell.X * scale;
            double y = cell.Y * scale;
            double w = cell.Width * scale;
            double h = cell.Height * scale;
            double paddingX = CellTextHorizontalPadding * scale;
            double paddingY = CellTextVerticalPadding * scale;
            double availableWidth = Math.Max(1D, w - (paddingX * 2D));
            double availableHeight = Math.Max(1D, h - (paddingY * 2D));
            double fontSize = ResolveCellFontSize(cell.Style, scale);
            double minimumFontSize = Math.Max(1D, scale);
            double rotationDegrees = ResolveExcelTextRotationDegrees(cell.Style.TextRotation, snapshot, cell, diagnostics);
            bool rotated = Math.Abs(rotationDegrees) > 0.0001D;
            bool richTextSupported = IsRichTextRenderingSupported(cell, rotated);
            OfficeTextBlockLayout layout = OfficeTextLayoutEngine.LayoutTextBlock(
                cell.Text,
                fontSize,
                rotated ? Math.Max(availableWidth, availableHeight) : availableWidth,
                rotated ? Math.Max(availableWidth, availableHeight) : availableHeight,
                CellTextLineHeightFactor,
                minimumFontSize,
                textMeasureCanvas.MeasureText,
                wrap: cell.Style.WrapText,
                forceSingleLine: rotated,
                shrinkToFit: cell.Style.ShrinkToFit);
            if (layout.Lines.Count == 0) {
                return;
            }

            if (cell.RichTextRuns.Count > 0) {
                if (richTextSupported && TryAppendSvgRichText(builder, cell, options, x, y, w, h, paddingX, paddingY, availableWidth, availableHeight, rotationDegrees, textMeasureCanvas.MeasureText, out OfficeRichTextBlockLayout richLayout)) {
                    AddTextClippingDiagnosticIfNeeded(richLayout, snapshot, cell, diagnostics);
                    if (rotated) {
                        AddRotatedTextApproximationDiagnostic(snapshot, cell, diagnostics);
                    }

                    return;
                }

                AddRichTextLayoutApproximationDiagnostic(snapshot, cell, diagnostics);
            }

            AddTextClippingDiagnosticIfNeeded(layout, snapshot, cell, diagnostics);
            OfficeColor color = ResolveCellTextColor(cell, options);
            OfficeTextAlignment alignment = ResolveTextAlignment(cell.Style.HorizontalAlignment);
            string clipId = "xl-text-" + cell.Row.ToString(System.Globalization.CultureInfo.InvariantCulture) + "-" + cell.Column.ToString(System.Globalization.CultureInfo.InvariantCulture);

            builder.AppendRectClipPathDefinition(clipId, x, y, w, h);
            builder.Append("<g").AppendClipPathReference(clipId).Append(">");
            if (rotated) {
                AddRotatedTextApproximationDiagnostic(snapshot, cell, diagnostics);
                OfficeTextLine line = layout.Lines[0];
                double centerX = x + (w / 2D);
                double centerY = y + (h / 2D);
                double textTop = centerY - (layout.FontSize / 2D);
                double baseline = textTop + (layout.FontSize * 0.84D);
                builder.AppendSvgTextElement(
                    line.Text,
                    centerX,
                    baseline,
                    layout.LineHeight,
                    color,
                    ResolveSvgFontFamily(cell.Style),
                    layout.FontSize,
                    OfficeTextAlignment.Center,
                    cell.Style.Bold,
                    cell.Style.Italic,
                    ShouldUnderlineText(cell, options),
                    rotationDegrees,
                    centerX,
                    centerY);
                builder.Append("</g>");
                return;
            }

            builder.AppendSvgTextBlock(
                layout,
                x + paddingX,
                y + paddingY,
                availableWidth,
                availableHeight,
                color,
                ResolveSvgFontFamily(cell.Style),
                alignment,
                ResolveTextVerticalAlignment(cell.Style.VerticalAlignment),
                cell.Style.Bold,
                cell.Style.Italic,
                ShouldUnderlineText(cell, options));

            builder.Append("</g>");
        }

        private static bool TryDrawRasterRichText(
            OfficeRasterCanvas canvas,
            ExcelVisualCell cell,
            ExcelImageExportOptions options,
            double scale,
            double x,
            double y,
            double w,
            double h,
            double paddingX,
            double paddingY,
            double availableWidth,
            double availableHeight,
            double rotationDegrees,
            out OfficeRichTextBlockLayout layout) {
            bool rotated = Math.Abs(rotationDegrees) > 0.0001D;
            if (!TryBuildRichTextLayout(cell, options, scale, availableWidth, availableHeight, rotationDegrees, canvas.MeasureText, out layout)) {
                return false;
            }

            double centerX = x + (w / 2D);
            double centerY = y + (h / 2D);
            OfficeTextAlignment alignment = rotated ? OfficeTextAlignment.Center : ResolveTextAlignment(cell.Style.HorizontalAlignment);
            double layoutWidth = rotated ? Math.Max(availableWidth, availableHeight) : availableWidth;
            double left = rotated ? centerX - (layoutWidth / 2D) : x + paddingX;
            double top = rotated
                ? centerY - (layout.Height / 2D)
                : OfficeTextPlacement.ResolveTop(y + paddingY, availableHeight, layout.Height, ResolveTextVerticalAlignment(cell.Style.VerticalAlignment));
            for (int lineIndex = 0; lineIndex < layout.Lines.Count; lineIndex++) {
                OfficeRichTextLine line = layout.Lines[lineIndex];
                if (line.Segments.Count == 0) {
                    continue;
                }

                double lineTop = top + (lineIndex * layout.LineHeight);
                double lineFontSize = Math.Max(1D, line.FontSize);
                double textTop = lineTop + Math.Max(0D, (layout.LineHeight - lineFontSize) / 2D);
                double baseline = textTop + (lineFontSize * 0.84D);
                double cursor = OfficeTextPlacement.ResolveLineLeft(left, layoutWidth, line.Width, alignment);
                for (int segmentIndex = 0; segmentIndex < line.Segments.Count; segmentIndex++) {
                    OfficeRichTextSegment segment = line.Segments[segmentIndex];
                    double runTop = baseline - (segment.FontSize * 0.84D);
                    canvas.DrawTextLine(segment.Text, cursor, runTop, segment.FontSize, segment.Color, segment.Bold, segment.Italic, OfficeTextAlignment.Left, rotationDegrees, centerX, centerY, segment.Underline, segment.Strikethrough);
                    cursor += segment.Width;
                }
            }

            return true;
        }

        private static bool TryAppendSvgRichText(
            StringBuilder builder,
            ExcelVisualCell cell,
            ExcelImageExportOptions options,
            double x,
            double y,
            double w,
            double h,
            double paddingX,
            double paddingY,
            double availableWidth,
            double availableHeight,
            double rotationDegrees,
            Func<string?, double, double> measure,
            out OfficeRichTextBlockLayout layout) {
            bool rotated = Math.Abs(rotationDegrees) > 0.0001D;
            if (!TryBuildRichTextLayout(cell, options, options.Scale, availableWidth, availableHeight, rotationDegrees, measure, out layout)) {
                return false;
            }

            double centerX = x + (w / 2D);
            double centerY = y + (h / 2D);
            OfficeTextAlignment alignment = rotated ? OfficeTextAlignment.Center : ResolveTextAlignment(cell.Style.HorizontalAlignment);
            double layoutWidth = rotated ? Math.Max(availableWidth, availableHeight) : availableWidth;
            double left = rotated ? centerX - (layoutWidth / 2D) : x + paddingX;
            double top = rotated
                ? centerY - (layout.Height / 2D)
                : OfficeTextPlacement.ResolveTop(y + paddingY, availableHeight, layout.Height, ResolveTextVerticalAlignment(cell.Style.VerticalAlignment));
            string clipId = "xl-text-" + cell.Row.ToString(System.Globalization.CultureInfo.InvariantCulture) + "-" + cell.Column.ToString(System.Globalization.CultureInfo.InvariantCulture);

            builder.AppendRectClipPathDefinition(clipId, x, y, w, h);
            builder.Append("<g").AppendClipPathReference(clipId);
            if (rotated) {
                builder.AppendRotateTransformAttribute(rotationDegrees, centerX, centerY);
            }

            builder.Append(">");
            for (int lineIndex = 0; lineIndex < layout.Lines.Count; lineIndex++) {
                OfficeRichTextLine line = layout.Lines[lineIndex];
                if (line.Segments.Count == 0) {
                    continue;
                }

                double lineTop = top + (lineIndex * layout.LineHeight);
                double lineFontSize = Math.Max(1D, line.FontSize);
                double textTop = lineTop + Math.Max(0D, (layout.LineHeight - lineFontSize) / 2D);
                double baseline = textTop + (lineFontSize * 0.84D);
                double cursor = OfficeTextPlacement.ResolveLineLeft(left, layoutWidth, line.Width, alignment);
                for (int segmentIndex = 0; segmentIndex < line.Segments.Count; segmentIndex++) {
                    OfficeRichTextSegment segment = line.Segments[segmentIndex];
                    builder.AppendSvgRichTextSegment(segment, cursor, baseline);
                    cursor += segment.Width;
                }
            }

            builder.Append("</g>");
            return true;
        }

        private static bool TryBuildRichTextLayout(
            ExcelVisualCell cell,
            ExcelImageExportOptions options,
            double scale,
            double availableWidth,
            double availableHeight,
            double rotationDegrees,
            Func<string?, double, double> measure,
            out OfficeRichTextBlockLayout layout) {
            var runs = new List<OfficeRichTextRun>(cell.RichTextRuns.Count);
            OfficeColor fallbackColor = ResolveCellTextColor(cell, options);
            bool fallbackUnderline = ShouldUnderlineText(cell, options);
            for (int i = 0; i < cell.RichTextRuns.Count; i++) {
                ExcelVisualTextRun run = cell.RichTextRuns[i];
                if (string.IsNullOrEmpty(run.Text)) {
                    continue;
                }

                double fontSize = ResolveRunFontSize(run, cell.Style, scale);
                OfficeColor color = ResolveArgb(run.FontColorArgb) ?? fallbackColor;
                bool bold = cell.Style.Bold || run.Bold;
                bool italic = cell.Style.Italic || run.Italic;
                bool underline = fallbackUnderline || run.Underline;
                runs.Add(new OfficeRichTextRun(run.Text, fontSize, color, bold, italic, underline, ResolveRunFontFamily(run, cell.Style)));
            }

            if (runs.Count == 0) {
                layout = new OfficeRichTextBlockLayout(Array.Empty<OfficeRichTextLine>(), 0D, 0D, 0D);
                return false;
            }

            bool rotated = Math.Abs(rotationDegrees) > 0.0001D;
            double estimatedLineHeight = Math.Ceiling(ResolveMaxRichTextRunFontSize(runs) * CellTextLineHeightFactor);
            double layoutWidth = rotated
                ? ResolveRotatedTextWidthLimit(availableWidth, availableHeight, estimatedLineHeight, rotationDegrees)
                : availableWidth;
            double layoutHeight = rotated ? Math.Max(availableWidth, availableHeight) : availableHeight;
            layout = OfficeTextLayoutEngine.LayoutRichTextBlock(
                runs,
                layoutWidth,
                layoutHeight,
                CellTextLineHeightFactor,
                measure,
                wrap: cell.Style.WrapText && !rotated,
                shrinkToFit: cell.Style.ShrinkToFit || rotated,
                minimumFontSize: Math.Max(1D, scale));
            return layout.Lines.Count > 0;
        }

        private static bool IsRichTextRenderingSupported(ExcelVisualCell cell, bool rotated) {
            return true;
        }

        private static double ResolveMaxRichTextRunFontSize(IReadOnlyList<OfficeRichTextRun> runs) {
            double max = 1D;
            for (int i = 0; i < runs.Count; i++) {
                max = Math.Max(max, runs[i].FontSize);
            }

            return max;
        }

        private static double ResolveRotatedTextWidthLimit(double availableWidth, double availableHeight, double lineHeight, double rotationDegrees) {
            double width = Math.Max(1D, availableWidth);
            double height = Math.Max(1D, availableHeight);
            double radians = Math.Abs(rotationDegrees) * Math.PI / 180D;
            double cos = Math.Abs(Math.Cos(radians));
            double sin = Math.Abs(Math.Sin(radians));
            double estimatedHeight = Math.Max(1D, lineHeight);
            double limit = Math.Max(width, height);

            if (cos > 0.000001D) {
                limit = Math.Min(limit, (width - (estimatedHeight * sin)) / cos);
            }

            if (sin > 0.000001D) {
                limit = Math.Min(limit, (height - (estimatedHeight * cos)) / sin);
            }

            if (double.IsNaN(limit) || double.IsInfinity(limit)) {
                return Math.Max(width, height);
            }

            return Math.Max(1D, limit);
        }

        private static double ResolveRunFontSize(ExcelVisualTextRun run, ExcelCellStyleSnapshot style, double scale) {
            double fontSize = run.FontSize.GetValueOrDefault(style.FontSize.GetValueOrDefault(CellTextFontSize));
            if (fontSize <= 0D || double.IsNaN(fontSize) || double.IsInfinity(fontSize)) {
                fontSize = CellTextFontSize;
            }

            return fontSize * scale;
        }

        private static string ResolveRunFontFamily(ExcelVisualTextRun run, ExcelCellStyleSnapshot style) {
            string? fontName = string.IsNullOrWhiteSpace(run.FontName) ? style.FontName : run.FontName;
            return string.IsNullOrWhiteSpace(fontName) ? "Arial, sans-serif" : fontName! + ", Arial, sans-serif";
        }

        private static double Measure(string? text, double fontSize, Func<string?, double, double> measure) =>
            string.IsNullOrEmpty(text) ? 0D : measure(text, fontSize);

        private static double ResolveCellFontSize(ExcelCellStyleSnapshot style, double scale) {
            double fontSize = style.FontSize.GetValueOrDefault(CellTextFontSize);
            if (fontSize <= 0D || double.IsNaN(fontSize) || double.IsInfinity(fontSize)) {
                fontSize = CellTextFontSize;
            }

            return fontSize * scale;
        }

        private static string ResolveSvgFontFamily(ExcelCellStyleSnapshot style) =>
            string.IsNullOrWhiteSpace(style.FontName) ? "Arial, sans-serif" : style.FontName! + ", Arial, sans-serif";

        private static OfficeColor ResolveCellTextColor(ExcelVisualCell cell, ExcelImageExportOptions options) {
            OfficeColor? explicitColor = ResolveArgb(cell.Style.FontColorArgb);
            if (explicitColor.HasValue) {
                return explicitColor.Value;
            }

            if (ShouldUseHyperlinkHint(cell, options)) {
                return HyperlinkHintColor;
            }

            return OfficeColor.Black;
        }

        private static bool ShouldUnderlineText(ExcelVisualCell cell, ExcelImageExportOptions options) =>
            cell.Style.Underline || ShouldUseHyperlinkHint(cell, options);

        private static bool ShouldUseHyperlinkHint(ExcelVisualCell cell, ExcelImageExportOptions options) =>
            options.ShowHyperlinkHints && cell.Hyperlink != null && !cell.Style.Underline;

        private static double ResolveExcelTextRotationDegrees(int? textRotation, ExcelRangeVisualSnapshot snapshot, ExcelVisualCell cell, List<OfficeImageExportDiagnostic>? diagnostics) {
            if (!textRotation.HasValue || textRotation.Value == 0) {
                return 0D;
            }

            int value = textRotation.Value;
            if (value == 255) {
                diagnostics?.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.CellStackedTextRotationUnsupported,
                    "Stacked vertical text rotation is not yet rendered exactly during image export; the cell is rendered without stacked text.",
                    GetCellDiagnosticSource(snapshot, cell)));
                return 0D;
            }

            if (value < 0 || value > 180) {
                diagnostics?.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.CellTextRotationUnsupported,
                    "The cell uses an unsupported text rotation value and is rendered without rotation.",
                    GetCellDiagnosticSource(snapshot, cell)));
                return 0D;
            }

            return value <= 90 ? -value : value - 90D;
        }

        private static void AddRotatedTextApproximationDiagnostic(ExcelRangeVisualSnapshot snapshot, ExcelVisualCell cell, List<OfficeImageExportDiagnostic>? diagnostics) {
            diagnostics?.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                ExcelImageExportDiagnosticCodes.CellTextRotationApproximation,
                "Rotated text was rendered using the shared drawing engine, but Excel baseline, anchoring, and stacked text behavior are still approximate.",
                GetCellDiagnosticSource(snapshot, cell)));
        }

        private static void AddTextClippingDiagnosticIfNeeded(OfficeTextBlockLayout layout, ExcelRangeVisualSnapshot snapshot, ExcelVisualCell cell, List<OfficeImageExportDiagnostic>? diagnostics) {
            if (!layout.Clipped) {
                return;
            }

            diagnostics?.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                ExcelImageExportDiagnosticCodes.CellTextClipped,
                "Cell text was clipped or ellipsized during image export because it does not fit the rendered cell bounds.",
                GetCellDiagnosticSource(snapshot, cell)));
        }

        private static void AddTextClippingDiagnosticIfNeeded(OfficeRichTextBlockLayout layout, ExcelRangeVisualSnapshot snapshot, ExcelVisualCell cell, List<OfficeImageExportDiagnostic>? diagnostics) {
            if (!layout.Clipped) {
                return;
            }

            diagnostics?.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                ExcelImageExportDiagnosticCodes.CellTextClipped,
                "Cell rich text was clipped or ellipsized during image export because it does not fit the rendered cell bounds.",
                GetCellDiagnosticSource(snapshot, cell)));
        }

        private static void AddRichTextLayoutApproximationDiagnostic(ExcelRangeVisualSnapshot snapshot, ExcelVisualCell cell, List<OfficeImageExportDiagnostic>? diagnostics) {
            diagnostics?.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                ExcelImageExportDiagnosticCodes.CellRichTextLayoutApproximation,
                "Cell rich text runs were detected, but this layout path cannot render the runs exactly yet; the cell was rendered as plain styled text.",
                GetCellDiagnosticSource(snapshot, cell)));
        }

        private static string GetCellDiagnosticSource(ExcelRangeVisualSnapshot snapshot, ExcelVisualCell cell) =>
            snapshot.SheetName + "!" + A1.ColumnIndexToLetters(cell.Column) + cell.Row.ToString(System.Globalization.CultureInfo.InvariantCulture);

    }
}
