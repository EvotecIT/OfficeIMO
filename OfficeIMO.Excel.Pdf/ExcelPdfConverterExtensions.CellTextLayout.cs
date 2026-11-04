using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        /// <summary>
        /// Horizontal cell padding of the worksheet canvas, in unscaled points: three
        /// pixels, so text and numbers keep clear of the gridlines as they do in Excel.
        /// </summary>
        private const double WorksheetCellPaddingX = 2.25D;

        /// <summary>
        /// Measures cell text with the font programs that the PDF writer paints with, so
        /// row heights, overflow and "####" decisions agree with the rendered glyphs.
        /// </summary>
        private sealed class WorksheetTextMeasurer {
            private readonly Func<string?, double, string?, OfficeFontStyle, double> _measure;
            private readonly string? _defaultFamily;

            public WorksheetTextMeasurer(PdfCore.PdfCreateOptions pdfOptions, string? defaultFamily, double defaultFontSize, double defaultRowHeight, double workbookFontSize) {
                _measure = PdfCore.PdfWriter.CreateDrawingTextMeasure(pdfOptions);
                _defaultFamily = string.IsNullOrWhiteSpace(defaultFamily) ? null : defaultFamily;
                DefaultFontSize = defaultFontSize > 0D ? defaultFontSize : 11D;
                DefaultRowHeight = defaultRowHeight > 0D ? defaultRowHeight : 15D;
                WorkbookFontSize = workbookFontSize > 0D ? workbookFontSize : 11D;
            }

            /// <summary>Font size used for cells without an authored size.</summary>
            public double DefaultFontSize { get; }

            /// <summary>Worksheet default row height in points.</summary>
            public double DefaultRowHeight { get; }

            /// <summary>Size of the workbook's Normal style font, which the default row height belongs to.</summary>
            public double WorkbookFontSize { get; }

            public double GetFontSize(ExcelCellStyleInfo? style) =>
                style?.FontSize is OfficeLength size && size > OfficeLength.Zero ? size.ToPoints() : DefaultFontSize;

            /// <summary>Excel's single-line row height for a font size, scaled from the default row.</summary>
            public double GetLineHeight(double fontSize) => DefaultRowHeight * fontSize / WorkbookFontSize;

            public double Measure(string text, ExcelCellStyleInfo? style) {
                if (string.IsNullOrEmpty(text)) return 0D;
                double size = GetFontSize(style);
                string? family = string.IsNullOrWhiteSpace(style?.FontName) ? _defaultFamily : style!.FontName;
                OfficeFontStyle fontStyle = (style?.Bold == true ? OfficeFontStyle.Bold : OfficeFontStyle.Regular)
                    | (style?.Italic == true ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);
                double width = _measure(text, size, family, fontStyle);
                return double.IsNaN(width) || double.IsInfinity(width) || width < 0D ? text.Length * size * 0.6D : width;
            }

            /// <summary>Counts the lines Excel's word wrap produces for a width in points.</summary>
            public int CountWrappedLines(string text, ExcelCellStyleInfo? style, double width) {
                if (string.IsNullOrEmpty(text)) return 1;
                int lines = 0;
                foreach (string paragraph in text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n')) {
                    lines++;
                    string[] words = paragraph.Split(' ');
                    double lineWidth = 0D;
                    double space = Measure(" ", style);
                    bool lineHasWord = false;
                    foreach (string word in words) {
                        double wordWidth = Measure(word, style);
                        if (!lineHasWord) {
                            lineWidth = wordWidth;
                            lineHasWord = true;
                        } else if (lineWidth + space + wordWidth <= width) {
                            lineWidth += space + wordWidth;
                        } else {
                            lines++;
                            lineWidth = wordWidth;
                        }

                        // A word longer than the cell breaks between characters.
                        while (lineWidth > width && width > 0D && wordWidth > width) {
                            lines++;
                            lineWidth -= width;
                            wordWidth -= width;
                        }
                    }
                }

                return Math.Max(1, lines);
            }
        }

        /// <summary>
        /// Prepares worksheet text layout after fonts are registered: a shared measurer
        /// and Excel's automatic heights for rows that store no explicit height.
        /// </summary>
        private static void ApplyWorksheetTextLayout(
            ExcelDocument document,
            PdfCore.PdfCreateOptions pdfOptions,
            ExcelPdfExportOptions options,
            IReadOnlyList<WorksheetPdfExportPlan> plans) {
            if (options.WorksheetLayout != ExcelPdfWorksheetLayoutMode.WorksheetCanvas) return;
            OfficeFontInfo? defaultFont = document.GetWorkbookDefaultFontInfo();
            foreach (WorksheetPdfExportPlan plan in plans) {
                var measurer = new WorksheetTextMeasurer(
                    pdfOptions,
                    defaultFont?.FamilyName,
                    GetDefaultTableFontSize(options),
                    plan.Geometry.DefaultRowHeight,
                    defaultFont?.Size ?? 11D);
                plan.TextMeasurer = measurer;
                if (options.UseWorksheetRowHeights && options.UseWorksheetCellStyles) {
                    plan.Geometry.AutoRowHeights = CalculateAutomaticRowHeights(plan, measurer, options);
                }
            }
        }

        private static Dictionary<int, double>? CalculateAutomaticRowHeights(WorksheetPdfExportPlan plan, WorksheetTextMeasurer measurer, ExcelPdfExportOptions options) {
            SheetExportData data = plan.ExportData;
            string?[,]? references = data.CellReferences;
            ExcelCellStyleInfo?[,]? styles = data.Styles;
            if (references == null || styles == null) return null;
            int rows = Math.Min(plan.ExportedRows, data.Values.GetLength(0));
            int columns = data.Values.GetLength(1);
            Dictionary<int, double>? heights = null;
            var columnWidths = new double[columns];
            for (int column = 0; column < columns; column++) columnWidths[column] = GetExportedColumnWidthPoints(plan, column);
            for (int row = 0; row < rows; row++) {
                options.CancellationToken.ThrowIfCancellationRequested();
                int sourceRow = GetOriginalRowNumber(references, row);
                if (sourceRow <= 0 || HasStoredRowHeight(plan.Geometry, sourceRow)) continue;
                double height = 0D;
                for (int column = 0; column < columns; column++) {
                    ExcelCellStyleInfo? style = styles[row, column];
                    if (style == null || style.TextRotation is int rotation && rotation != 0) continue;
                    // Like Excel, a merge across rows never sizes a row, and merged text is
                    // not wrapped for sizing; its font still sets the line height.
                    MergeRegion? region = data.MergedCells?.GetRegion(row, column);
                    if (region != null && region.Span.RowSpan > 1) continue;
                    if (region != null && (region.Row != row || region.Column != column)) continue;

                    double fontSize = measurer.GetFontSize(style);
                    double cellHeight = measurer.GetLineHeight(fontSize);
                    if (style.WrapText && region == null && data.Values[row, column] is { } value) {
                        string text = FormatCellValue(value, style, string.Empty, ExcelDateSystem.NineteenHundred);
                        double width = Math.Max(1D, columnWidths[column] - 2D * WorksheetCellPaddingX);
                        cellHeight *= measurer.CountWrappedLines(text, style, width);
                    }

                    height = Math.Max(height, cellHeight);
                }

                if (height > plan.Geometry.DefaultRowHeight + 0.01D) {
                    // Excel stores row heights on the 0.75 pt (one pixel) grid.
                    heights ??= new Dictionary<int, double>();
                    heights[sourceRow] = Math.Ceiling(height / 0.75D) * 0.75D;
                }
            }

            return heights;
        }

        private static bool HasStoredRowHeight(WorksheetGeometryData geometry, int sourceRow) {
            for (int index = geometry.Rows.Count - 1; index >= 0; index--) {
                ExcelRowInfo definition = geometry.Rows[index];
                if (definition.Index == sourceRow) {
                    return definition.StoredHeight is double height && height > 0D;
                }
            }

            return false;
        }
    }
}
