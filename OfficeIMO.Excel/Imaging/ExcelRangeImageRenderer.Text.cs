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
            IReadOnlyDictionary<string, ExcelVisualCell> cellsByAddress,
            List<OfficeImageExportDiagnostic>? diagnostics) {
            if (string.IsNullOrEmpty(cell.Text)) {
                return;
            }

            CellTextViewport viewport = ResolveCellTextViewport(cell, snapshot, scale, cellsByAddress);
            double x = viewport.X;
            double y = viewport.Y;
            double w = viewport.Width;
            double h = viewport.Height;
            double paddingX = CellTextHorizontalPadding * scale;
            double paddingY = CellTextVerticalPadding * scale;
            double availableWidth = Math.Max(1D, w - (paddingX * 2D));
            double availableHeight = Math.Max(1D, h - (paddingY * 2D));
            double fontSize = ResolveCellFontSize(cell.Style, scale);
            double minimumFontSize = Math.Max(1D, scale);
            bool stacked = IsStackedTextRotation(cell.Style.TextRotation);
            double rotationDegrees = stacked ? 0D : ResolveExcelTextRotationDegrees(cell.Style.TextRotation, snapshot, cell, diagnostics);
            bool rotated = Math.Abs(rotationDegrees) > 0.0001D;
            bool richTextSupported = IsRichTextRenderingSupported(cell, rotated);
            string fontFamily = ResolveCellFontFamily(cell.Style);
            OfficeTextBlockRenderPlan plan = CreateCellTextRenderPlan(
                cell,
                x + paddingX,
                y + paddingY,
                availableWidth,
                availableHeight,
                fontSize,
                minimumFontSize,
                rotationDegrees,
                stacked,
                (text, size) => canvas.MeasureText(text, size, fontFamily));
            OfficeTextBlockLayout layout = plan.Layout;
            if (layout.Lines.Count == 0) {
                return;
            }

            using (canvas.PushClipRectangle(x, y, w, h)) {
                if (cell.RichTextRuns.Count > 0) {
                    if (richTextSupported && TryDrawRasterRichText(canvas, cell, options, scale, x, y, w, h, paddingX, paddingY, availableWidth, availableHeight, rotationDegrees, stacked, out OfficeRichTextBlockLayout richLayout)) {
                        AddRichTextFontFamilyFallbackDiagnostics(snapshot, cell, diagnostics);
                        AddTextClippingDiagnosticIfNeeded(richLayout, snapshot, cell, diagnostics);
                        if (rotated || stacked) {
                            AddRotatedTextApproximationDiagnostic(snapshot, cell, diagnostics);
                        }

                        return;
                    }

                    AddRichTextLayoutApproximationDiagnostic(snapshot, cell, diagnostics);
                }

                AddCellFontFamilyFallbackDiagnosticIfNeeded(snapshot, cell, cell.Style.FontName, diagnostics);
                AddTextClippingDiagnosticIfNeeded(layout, snapshot, cell, diagnostics);
                if (stacked) {
                    AddRotatedTextApproximationDiagnostic(snapshot, cell, diagnostics);
                    OfficeTextBlockRenderer.DrawRasterTextBlock(
                        canvas,
                        plan.Layout,
                        plan.Left,
                        plan.Top,
                        plan.Width,
                        plan.Height,
                        ResolveCellTextColor(cell, options),
                        plan.HorizontalAlignment,
                        plan.VerticalAlignment,
                        cell.Style.Bold,
                        cell.Style.Italic,
                        ShouldUnderlineText(cell, options),
                        fontFamily: fontFamily);
                    return;
                }

                OfficeColor color = ResolveCellTextColor(cell, options);
                bool bold = cell.Style.Bold;
                bool italic = cell.Style.Italic;
                bool underline = ShouldUnderlineText(cell, options);
                if (rotated) {
                    AddRotatedTextApproximationDiagnostic(snapshot, cell, diagnostics);
                    OfficeTextLine line = layout.Lines[0];
                    double centerX = x + (w / 2D);
                    double centerY = y + (h / 2D);
                    double textTop = centerY - (layout.FontSize / 2D);
                    canvas.DrawTextLine(line.Text, centerX, textTop, layout.FontSize, color, bold, italic, OfficeTextAlignment.Center, rotationDegrees, centerX, centerY, fontFamily: fontFamily);
                    return;
                }

                OfficeTextBlockRenderer.DrawRasterTextBlock(
                    canvas,
                    plan.Layout,
                    plan.Left,
                    plan.Top,
                    plan.Width,
                    plan.Height,
                    color,
                    plan.HorizontalAlignment,
                    plan.VerticalAlignment,
                    bold,
                    italic,
                    underline,
                    fontFamily: fontFamily);
            }
        }

        private static void AppendSvgCellText(
            StringBuilder builder,
            ExcelVisualCell cell,
            ExcelRangeVisualSnapshot snapshot,
            ExcelImageExportOptions options,
            OfficeRasterCanvas textMeasureCanvas,
            IReadOnlyDictionary<string, ExcelVisualCell> cellsByAddress,
            List<OfficeImageExportDiagnostic>? diagnostics) {
            if (string.IsNullOrEmpty(cell.Text)) {
                return;
            }

            double scale = options.Scale;
            CellTextViewport viewport = ResolveCellTextViewport(cell, snapshot, scale, cellsByAddress);
            double x = viewport.X;
            double y = viewport.Y;
            double w = viewport.Width;
            double h = viewport.Height;
            double paddingX = CellTextHorizontalPadding * scale;
            double paddingY = CellTextVerticalPadding * scale;
            double availableWidth = Math.Max(1D, w - (paddingX * 2D));
            double availableHeight = Math.Max(1D, h - (paddingY * 2D));
            double fontSize = ResolveCellFontSize(cell.Style, scale);
            double minimumFontSize = Math.Max(1D, scale);
            bool stacked = IsStackedTextRotation(cell.Style.TextRotation);
            double rotationDegrees = stacked ? 0D : ResolveExcelTextRotationDegrees(cell.Style.TextRotation, snapshot, cell, diagnostics);
            bool rotated = Math.Abs(rotationDegrees) > 0.0001D;
            bool richTextSupported = IsRichTextRenderingSupported(cell, rotated);
            string fontFamily = ResolveCellFontFamily(cell.Style);
            OfficeTextBlockRenderPlan plan = CreateCellTextRenderPlan(
                cell,
                x + paddingX,
                y + paddingY,
                availableWidth,
                availableHeight,
                fontSize,
                minimumFontSize,
                rotationDegrees,
                stacked,
                (text, size) => textMeasureCanvas.MeasureText(text, size, fontFamily));
            OfficeTextBlockLayout layout = plan.Layout;
            if (layout.Lines.Count == 0) {
                return;
            }

            if (cell.RichTextRuns.Count > 0) {
                if (richTextSupported && TryAppendSvgRichText(builder, cell, options, x, y, w, h, paddingX, paddingY, availableWidth, availableHeight, rotationDegrees, stacked, (text, size, family) => textMeasureCanvas.MeasureText(text, size, family), out OfficeRichTextBlockLayout richLayout)) {
                    AddRichTextFontFamilyFallbackDiagnostics(snapshot, cell, diagnostics);
                    AddTextClippingDiagnosticIfNeeded(richLayout, snapshot, cell, diagnostics);
                    if (rotated || stacked) {
                        AddRotatedTextApproximationDiagnostic(snapshot, cell, diagnostics);
                    }

                    return;
                }

                AddRichTextLayoutApproximationDiagnostic(snapshot, cell, diagnostics);
            }

            AddCellFontFamilyFallbackDiagnosticIfNeeded(snapshot, cell, cell.Style.FontName, diagnostics);
            AddTextClippingDiagnosticIfNeeded(layout, snapshot, cell, diagnostics);
            OfficeColor color = ResolveCellTextColor(cell, options);
            string clipId = "xl-text-" + cell.Row.ToString(System.Globalization.CultureInfo.InvariantCulture) + "-" + cell.Column.ToString(System.Globalization.CultureInfo.InvariantCulture);

            builder.AppendRectClipPathDefinition(clipId, x, y, w, h);
            builder.Append("<g").AppendClipPathReference(clipId).Append(">");
            if (stacked) {
                AddRotatedTextApproximationDiagnostic(snapshot, cell, diagnostics);
                builder.AppendSvgTextBlock(
                    plan.Layout,
                    plan.Left,
                    plan.Top,
                    plan.Width,
                    plan.Height,
                    color,
                    fontFamily,
                    plan.HorizontalAlignment,
                    plan.VerticalAlignment,
                    cell.Style.Bold,
                    cell.Style.Italic,
                    ShouldUnderlineText(cell, options));

                builder.Append("</g>");
                return;
            }

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
                    fontFamily,
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
                plan.Layout,
                plan.Left,
                plan.Top,
                plan.Width,
                plan.Height,
                color,
                fontFamily,
                plan.HorizontalAlignment,
                plan.VerticalAlignment,
                cell.Style.Bold,
                cell.Style.Italic,
                ShouldUnderlineText(cell, options));

            builder.Append("</g>");
        }

        private static OfficeTextBlockRenderPlan CreateCellTextRenderPlan(
            ExcelVisualCell cell,
            double left,
            double top,
            double availableWidth,
            double availableHeight,
            double fontSize,
            double minimumFontSize,
            double rotationDegrees,
            bool stacked,
            Func<string?, double, double> measure) {
            bool rotated = Math.Abs(rotationDegrees) > 0.0001D;
            OfficeTextVerticalAlignment verticalAlignment = rotated
                ? OfficeTextVerticalAlignment.Top
                : ResolveTextVerticalAlignment(cell.Style.VerticalAlignment);
            if (stacked) {
                return OfficeTextBlockRenderPlan.CreateStackedTextBlockFromRectangle(
                    cell.Text,
                    fontSize,
                    left,
                    top,
                    availableWidth,
                    availableHeight,
                    measure,
                    OfficeTextAlignment.Center,
                    verticalAlignment,
                    CellTextLineHeightFactor,
                    minimumFontSize,
                    shrinkToFit: cell.Style.ShrinkToFit);
            }

            double lineHeight = Math.Ceiling(fontSize * CellTextLineHeightFactor);
            double layoutWidth = rotated
                ? OfficeTextLayoutEngine.ResolveRotatedTextWidthLimit(availableWidth, availableHeight, lineHeight, rotationDegrees)
                : availableWidth;
            double layoutHeight = rotated ? Math.Max(availableWidth, availableHeight) : availableHeight;
            OfficeTextAlignment alignment = rotated
                ? OfficeTextAlignment.Center
                : ResolveCellTextAlignment(cell);
            return OfficeTextBlockRenderPlan.CreateTextBlockFromRectangle(
                cell.Text,
                fontSize,
                left,
                top,
                layoutWidth,
                layoutHeight,
                measure,
                alignment,
                verticalAlignment,
                CellTextLineHeightFactor,
                minimumFontSize,
                wrap: cell.Style.WrapText,
                forceSingleLine: rotated,
                shrinkToFit: cell.Style.ShrinkToFit,
                overflowBehavior: OfficeTextOverflowBehavior.Clip);
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
            bool stacked,
            out OfficeRichTextBlockLayout layout) {
            bool rotated = Math.Abs(rotationDegrees) > 0.0001D;
            if (!TryBuildRichTextLayout(cell, options, scale, availableWidth, availableHeight, rotationDegrees, stacked, (text, size, family) => canvas.MeasureText(text, size, family), out layout)) {
                return false;
            }

            double centerX = x + (w / 2D);
            double centerY = y + (h / 2D);
            OfficeTextAlignment alignment = (rotated || stacked) ? OfficeTextAlignment.Center : ResolveCellTextAlignment(cell);
            double layoutWidth = rotated ? Math.Max(availableWidth, availableHeight) : availableWidth;
            double left = rotated ? centerX - (layoutWidth / 2D) : x + paddingX;
            double top = rotated ? centerY - (layout.Height / 2D) : y + paddingY;
            double height = rotated ? layout.Height : availableHeight;
            OfficeTextBlockRenderer.DrawRasterRichTextBlock(
                canvas,
                layout,
                left,
                top,
                layoutWidth,
                height,
                alignment,
                rotated ? OfficeTextVerticalAlignment.Top : ResolveTextVerticalAlignment(cell.Style.VerticalAlignment),
                rotationDegrees,
                centerX,
                centerY);

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
            bool stacked,
            Func<string?, double, string?, double> measure,
            out OfficeRichTextBlockLayout layout) {
            bool rotated = Math.Abs(rotationDegrees) > 0.0001D;
            if (!TryBuildRichTextLayout(cell, options, options.Scale, availableWidth, availableHeight, rotationDegrees, stacked, measure, out layout)) {
                return false;
            }

            double centerX = x + (w / 2D);
            double centerY = y + (h / 2D);
            OfficeTextAlignment alignment = (rotated || stacked) ? OfficeTextAlignment.Center : ResolveCellTextAlignment(cell);
            double layoutWidth = rotated ? Math.Max(availableWidth, availableHeight) : availableWidth;
            double left = rotated ? centerX - (layoutWidth / 2D) : x + paddingX;
            double top = rotated ? centerY - (layout.Height / 2D) : y + paddingY;
            string clipId = "xl-text-" + cell.Row.ToString(System.Globalization.CultureInfo.InvariantCulture) + "-" + cell.Column.ToString(System.Globalization.CultureInfo.InvariantCulture);

            builder.AppendRectClipPathDefinition(clipId, x, y, w, h);
            builder.Append("<g").AppendClipPathReference(clipId);
            if (rotated) {
                builder.AppendRotateTransformAttribute(rotationDegrees, centerX, centerY);
            }

            builder.Append(">");
            builder.AppendSvgRichTextBlock(
                layout,
                left,
                top,
                layoutWidth,
                rotated ? layout.Height : availableHeight,
                alignment,
                rotated ? OfficeTextVerticalAlignment.Top : ResolveTextVerticalAlignment(cell.Style.VerticalAlignment));

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
            bool stacked,
            Func<string?, double, string?, double> measure,
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
            if (stacked) {
                layout = OfficeTextLayoutEngine.LayoutStackedRichTextBlock(
                    runs,
                    availableWidth,
                    availableHeight,
                    CellTextLineHeightFactor,
                    measure,
                    shrinkToFit: cell.Style.ShrinkToFit,
                    minimumFontSize: Math.Max(1D, scale));
                return layout.Lines.Count > 0;
            }

            double estimatedLineHeight = Math.Ceiling(ResolveMaxRichTextRunFontSize(runs) * CellTextLineHeightFactor);
            double layoutWidth = rotated
                ? OfficeTextLayoutEngine.ResolveRotatedTextWidthLimit(availableWidth, availableHeight, estimatedLineHeight, rotationDegrees)
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
                minimumFontSize: Math.Max(1D, scale),
                overflowBehavior: OfficeTextOverflowBehavior.Clip);
            return layout.Lines.Count > 0;
        }

        private static bool IsRichTextRenderingSupported(ExcelVisualCell cell, bool rotated) {
            return true;
        }

        private static CellTextViewport ResolveCellTextViewport(
            ExcelVisualCell cell,
            ExcelRangeVisualSnapshot snapshot,
            double scale,
            IReadOnlyDictionary<string, ExcelVisualCell> cellsByAddress) {
            double x = cell.X * scale;
            double y = cell.Y * scale;
            double width = cell.Width * scale;
            double height = cell.Height * scale;
            if (!CanCellTextSpillRight(cell, snapshot)) {
                return new CellTextViewport(x, y, width, height);
            }

            double unscaledRight = cell.X + cell.Width;
            for (int column = cell.Column + 1; column <= snapshot.LastColumn; column++) {
                if (!cellsByAddress.TryGetValue(Key(cell.Row, column), out ExcelVisualCell? neighbor) ||
                    !CanSpillThroughNeighbor(neighbor, snapshot)) {
                    break;
                }

                unscaledRight = neighbor.X + neighbor.Width;
            }

            return new CellTextViewport(x, y, Math.Max(width, (unscaledRight - cell.X) * scale), height);
        }

        private static bool CanCellTextSpillRight(ExcelVisualCell cell, ExcelRangeVisualSnapshot snapshot) {
            if (cell.CoveredByMerge ||
                cell.Style.WrapText ||
                cell.Style.ShrinkToFit ||
                IsStackedTextRotation(cell.Style.TextRotation) ||
                cell.Style.TextRotation.GetValueOrDefault() != 0 ||
                IsCellCoveredByDrawingLayer(cell, snapshot) ||
                cell.Text.IndexOf('\n') >= 0 ||
                cell.Text.IndexOf('\r') >= 0) {
                return false;
            }

            string? alignment = cell.Style.HorizontalAlignment;
            if (string.Equals(alignment, "left", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            return (string.IsNullOrWhiteSpace(alignment) || string.Equals(alignment, "general", StringComparison.OrdinalIgnoreCase)) &&
                cell.ValueKind == ExcelVisualCellValueKind.Text;
        }

        private static bool CanSpillThroughNeighbor(ExcelVisualCell neighbor, ExcelRangeVisualSnapshot snapshot) {
            return !neighbor.CoveredByMerge &&
                string.IsNullOrEmpty(neighbor.Text) &&
                neighbor.RichTextRuns.Count == 0 &&
                !IsCellCoveredByDrawingLayer(neighbor, snapshot);
        }

        private static bool IsCellCoveredByDrawingLayer(ExcelVisualCell cell, ExcelRangeVisualSnapshot snapshot) {
            for (int index = 0; index < snapshot.DrawingLayers.Count; index++) {
                if (TryGetDrawingLayerBounds(snapshot.DrawingLayers[index], out double x, out double y, out double width, out double height) &&
                    RectanglesIntersect(cell.X, cell.Y, cell.Width, cell.Height, x, y, width, height)) {
                    return true;
                }
            }

            return false;
        }

        private static bool TryGetDrawingLayerBounds(ExcelVisualDrawingLayer layer, out double x, out double y, out double width, out double height) {
            switch (layer.Kind) {
                case ExcelVisualDrawingLayerKind.DrawingObject when layer.DrawingObject != null:
                    x = layer.DrawingObject.X;
                    y = layer.DrawingObject.Y;
                    width = layer.DrawingObject.Width;
                    height = layer.DrawingObject.Height;
                    return width > 0D && height > 0D;
                case ExcelVisualDrawingLayerKind.Image when layer.Image != null:
                    x = layer.Image.X;
                    y = layer.Image.Y;
                    width = layer.Image.Width;
                    height = layer.Image.Height;
                    return width > 0D && height > 0D;
                case ExcelVisualDrawingLayerKind.Chart when layer.Chart != null:
                    x = layer.Chart.X;
                    y = layer.Chart.Y;
                    width = layer.Chart.Width;
                    height = layer.Chart.Height;
                    return width > 0D && height > 0D;
                case ExcelVisualDrawingLayerKind.CommentBody when layer.CommentBody != null:
                    x = layer.CommentBody.X;
                    y = layer.CommentBody.Y;
                    width = layer.CommentBody.Width;
                    height = layer.CommentBody.Height;
                    return width > 0D && height > 0D;
                default:
                    x = 0D;
                    y = 0D;
                    width = 0D;
                    height = 0D;
                    return false;
            }
        }

        private static bool RectanglesIntersect(double left, double top, double width, double height, double otherLeft, double otherTop, double otherWidth, double otherHeight) {
            if (width <= 0D || height <= 0D || otherWidth <= 0D || otherHeight <= 0D) {
                return false;
            }

            return left < otherLeft + otherWidth &&
                left + width > otherLeft &&
                top < otherTop + otherHeight &&
                top + height > otherTop;
        }

        private static OfficeTextAlignment ResolveCellTextAlignment(ExcelVisualCell cell) {
            string? alignment = cell.Style.HorizontalAlignment;
            if (string.Equals(alignment, "center", StringComparison.OrdinalIgnoreCase)) {
                return OfficeTextAlignment.Center;
            }

            if (string.Equals(alignment, "right", StringComparison.OrdinalIgnoreCase)) {
                return OfficeTextAlignment.Right;
            }

            if (string.Equals(alignment, "left", StringComparison.OrdinalIgnoreCase)) {
                return OfficeTextAlignment.Left;
            }

            if (!string.IsNullOrWhiteSpace(alignment) && !string.Equals(alignment, "general", StringComparison.OrdinalIgnoreCase)) {
                return ResolveTextAlignment(alignment);
            }

            return cell.ValueKind == ExcelVisualCellValueKind.Number || cell.ValueKind == ExcelVisualCellValueKind.Date
                ? OfficeTextAlignment.Right
                : cell.ValueKind == ExcelVisualCellValueKind.Boolean || cell.ValueKind == ExcelVisualCellValueKind.Error
                    ? OfficeTextAlignment.Center
                    : OfficeTextAlignment.Left;
        }

        private static double ResolveMaxRichTextRunFontSize(IReadOnlyList<OfficeRichTextRun> runs) {
            double max = 1D;
            for (int i = 0; i < runs.Count; i++) {
                max = Math.Max(max, runs[i].FontSize);
            }

            return max;
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

        private static double ResolveCellFontSize(ExcelCellStyleSnapshot style, double scale) {
            double fontSize = style.FontSize.GetValueOrDefault(CellTextFontSize);
            if (fontSize <= 0D || double.IsNaN(fontSize) || double.IsInfinity(fontSize)) {
                fontSize = CellTextFontSize;
            }

            return fontSize * scale;
        }

        private static string ResolveCellFontFamily(ExcelCellStyleSnapshot style) =>
            string.IsNullOrWhiteSpace(style.FontName) ? "Arial, sans-serif" : style.FontName! + ", Arial, sans-serif";

        private static void AddRichTextFontFamilyFallbackDiagnostics(ExcelRangeVisualSnapshot snapshot, ExcelVisualCell cell, List<OfficeImageExportDiagnostic>? diagnostics) {
            if (diagnostics == null || cell.RichTextRuns.Count == 0) {
                return;
            }

            var reported = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int index = 0; index < cell.RichTextRuns.Count; index++) {
                ExcelVisualTextRun run = cell.RichTextRuns[index];
                string? fontName = string.IsNullOrWhiteSpace(run.FontName) ? cell.Style.FontName : run.FontName;
                if (string.IsNullOrWhiteSpace(fontName) || !reported.Add(fontName!)) {
                    continue;
                }

                AddCellFontFamilyFallbackDiagnosticIfNeeded(snapshot, cell, fontName, diagnostics);
            }
        }

        private static void AddCellFontFamilyFallbackDiagnosticIfNeeded(
            ExcelRangeVisualSnapshot snapshot,
            ExcelVisualCell cell,
            string? fontName,
            List<OfficeImageExportDiagnostic>? diagnostics) {
            if (diagnostics == null || string.IsNullOrWhiteSpace(fontName) || OfficeTrueTypeFont.TryLoadFontFamily(fontName, out _) != null) {
                return;
            }

            diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                ExcelImageExportDiagnosticCodes.CellFontFamilyFallback,
                "Cell font family '" + fontName + "' could not be loaded exactly by the dependency-free image exporter; raster text metrics and image output used the shared fallback font path.",
                GetCellDiagnosticSource(snapshot, cell)));
        }

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
                "Cell text rotation was rendered using the shared drawing engine, but Excel baseline, anchoring, and stacked text behavior are still approximate.",
                GetCellDiagnosticSource(snapshot, cell)));
        }

        private static bool IsStackedTextRotation(int? textRotation) => textRotation == 255;

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

        private readonly struct CellTextViewport {
            internal CellTextViewport(double x, double y, double width, double height) {
                X = x;
                Y = y;
                Width = width;
                Height = height;
            }

            internal double X { get; }

            internal double Y { get; }

            internal double Width { get; }

            internal double Height { get; }
        }

    }
}
