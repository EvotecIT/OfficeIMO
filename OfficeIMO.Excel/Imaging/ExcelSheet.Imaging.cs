using System.IO;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Creates a format-neutral visual snapshot for a worksheet range or the used range.
        /// </summary>
        public ExcelRangeVisualSnapshot CreateVisualSnapshot(ExcelWorksheetImageExportOptions? options = null) {
            ExcelWorksheetImageExportOptions resolved = NormalizeWorksheetOptions(options);
            WorksheetImageRangeResolution range = ResolveWorksheetImageRange(resolved);
            return ExcelRangeVisualSnapshotBuilder.Build(this, range.Range, resolved, range.Diagnostics);
        }

        /// <summary>
        /// Exports a worksheet range or used range as PNG or SVG.
        /// </summary>
        public OfficeImageExportResult ExportImage(OfficeImageExportFormat format, ExcelWorksheetImageExportOptions? options = null) {
            ExcelWorksheetImageExportOptions resolved = NormalizeWorksheetOptions(options);
            WorksheetImageRangeResolution range = ResolveWorksheetImageRange(resolved);
            ExcelRangeVisualSnapshot snapshot = ExcelRangeVisualSnapshotBuilder.Build(this, range.Range, resolved, range.Diagnostics);
            return ExcelRangeImageRenderer.Render(snapshot, format, resolved);
        }

        /// <summary>
        /// Renders the worksheet used range to PNG bytes.
        /// </summary>
        public byte[] ToPng(ExcelWorksheetImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Png, options).Bytes;

        /// <summary>
        /// Renders the worksheet used range to SVG text.
        /// </summary>
        public string ToSvg(ExcelWorksheetImageExportOptions? options = null) =>
            Encoding.UTF8.GetString(ExportImage(OfficeImageExportFormat.Svg, options).Bytes);

        /// <summary>
        /// Saves the worksheet used range as PNG.
        /// </summary>
        public void SaveAsPng(string path, ExcelWorksheetImageExportOptions? options = null) =>
            WriteImageFile(path, ToPng(options));

        /// <summary>
        /// Saves the worksheet used range as SVG.
        /// </summary>
        public void SaveAsSvg(string path, ExcelWorksheetImageExportOptions? options = null) =>
            WriteImageFile(path, Encoding.UTF8.GetBytes(ToSvg(options)));

        /// <summary>
        /// Writes the worksheet used range as PNG to a stream.
        /// </summary>
        public void SaveAsPng(Stream stream, ExcelWorksheetImageExportOptions? options = null) =>
            WriteImageStream(stream, ToPng(options));

        /// <summary>
        /// Writes the worksheet used range as SVG to a stream.
        /// </summary>
        public void SaveAsSvg(Stream stream, ExcelWorksheetImageExportOptions? options = null) =>
            WriteImageStream(stream, Encoding.UTF8.GetBytes(ToSvg(options)));

        private static ExcelWorksheetImageExportOptions NormalizeWorksheetOptions(ExcelWorksheetImageExportOptions? options) {
            ExcelWorksheetImageExportOptions resolved = options ?? new ExcelWorksheetImageExportOptions();
            if (resolved.Scale <= 0D || double.IsNaN(resolved.Scale) || double.IsInfinity(resolved.Scale)) {
                throw new ArgumentOutOfRangeException(nameof(options), "Scale must be a finite positive number.");
            }

            return resolved;
        }

        private WorksheetImageRangeResolution ResolveWorksheetImageRange(ExcelWorksheetImageExportOptions options) {
            if (!string.IsNullOrWhiteSpace(options.Range)) {
                return new WorksheetImageRangeResolution(options.Range!, Array.Empty<OfficeImageExportDiagnostic>());
            }

            var diagnostics = new List<OfficeImageExportDiagnostic>();
            if (options.UsePrintArea) {
                string? printArea = GetPrintArea();
                string source = Name + "!_xlnm.Print_Area";
                if (string.IsNullOrWhiteSpace(printArea)) {
                    diagnostics.Add(new OfficeImageExportDiagnostic(
                        OfficeImageExportDiagnosticSeverity.Info,
                        ExcelImageExportDiagnosticCodes.PrintAreaMissing,
                        "Worksheet image export requested the print area, but no worksheet print area is configured; exporting the worksheet used range instead.",
                        source));
                } else if (SplitDefinedNameParts(printArea!).Skip(1).Any()) {
                    diagnostics.Add(new OfficeImageExportDiagnostic(
                        OfficeImageExportDiagnosticSeverity.Warning,
                        ExcelImageExportDiagnosticCodes.PrintAreaMultipleAreasUnsupported,
                        "Multi-area worksheet print areas are not supported by the image exporter; exporting the worksheet used range instead.",
                        source));
                } else if (TryNormalizeWorksheetImageRange(printArea!, out string? normalizedPrintArea)) {
                    return new WorksheetImageRangeResolution(normalizedPrintArea!, diagnostics);
                } else {
                    diagnostics.Add(new OfficeImageExportDiagnostic(
                        OfficeImageExportDiagnosticSeverity.Warning,
                        ExcelImageExportDiagnosticCodes.PrintAreaUnsupported,
                        "Worksheet print area could not be parsed as a supported A1 range; exporting the worksheet used range instead.",
                        source));
                }
            }

            return new WorksheetImageRangeResolution(ResolveWorksheetUsedImageRange(options), diagnostics);
        }

        private string ResolveWorksheetUsedImageRange(ExcelWorksheetImageExportOptions options) {
            string range = GetUsedRangeA1();
            if (!A1.TryParseRange(range, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                return range;
            }

            IReadOnlyList<ExcelColumnSnapshot> columns = GetColumnDefinitions();
            Dictionary<int, ExcelRowSnapshot> rows = GetRowDefinitions().ToDictionary(row => row.Index);
            if (options.IncludeImages) {
                foreach (ExcelImage image in Images) {
                    ExpandVisualAnchor(image.RowIndex, image.ColumnIndex, image.WidthPixels, image.HeightPixels, columns, rows, options, ref firstRow, ref firstColumn, ref lastRow, ref lastColumn);
                }
            }

            if (options.IncludeCharts) {
                foreach (ExcelChart chart in Charts) {
                    if (chart.TryGetSnapshot(out ExcelChartSnapshot snapshot)) {
                        ExpandVisualAnchor(snapshot.RowIndex, snapshot.ColumnIndex, snapshot.WidthPixels, snapshot.HeightPixels, columns, rows, options, ref firstRow, ref firstColumn, ref lastRow, ref lastColumn);
                    }
                }
            }

            return A1.CellReference(firstRow, firstColumn) + ":" + A1.CellReference(lastRow, lastColumn);
        }

        private static bool TryNormalizeWorksheetImageRange(string range, out string? normalizedRange) {
            string withoutSheet = StripSheetPrefix(range).Replace("$", string.Empty).Trim();
            if (A1.TryParseRange(withoutSheet, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                normalizedRange = A1.CellReference(firstRow, firstColumn) + ":" + A1.CellReference(lastRow, lastColumn);
                return true;
            }

            (int Row, int Col) cell = A1.ParseCellRef(withoutSheet);
            if (cell.Row > 0 && cell.Col > 0) {
                string reference = A1.CellReference(cell.Row, cell.Col);
                normalizedRange = reference + ":" + reference;
                return true;
            }

            normalizedRange = null;
            return false;
        }

        private static void ExpandVisualAnchor(
            int rowIndex,
            int columnIndex,
            int widthPixels,
            int heightPixels,
            IReadOnlyList<ExcelColumnSnapshot> columns,
            IReadOnlyDictionary<int, ExcelRowSnapshot> rows,
            ExcelImageExportOptions options,
            ref int firstRow,
            ref int firstColumn,
            ref int lastRow,
            ref int lastColumn) {
            if (rowIndex <= 0 || columnIndex <= 0) {
                return;
            }

            firstRow = Math.Min(firstRow, rowIndex);
            firstColumn = Math.Min(firstColumn, columnIndex);
            lastRow = Math.Max(lastRow, ResolveLastVisualRow(rowIndex, heightPixels, rows, options));
            lastColumn = Math.Max(lastColumn, ResolveLastVisualColumn(columnIndex, widthPixels, columns, options));
        }

        private static int ResolveLastVisualColumn(int startColumn, int widthPixels, IReadOnlyList<ExcelColumnSnapshot> columns, ExcelImageExportOptions options) {
            double remaining = Math.Max(1D, widthPixels);
            int column = startColumn;
            while (column < 16384) {
                remaining -= ResolveColumnWidth(columns.FirstOrDefault(item => column >= item.StartIndex && column <= item.EndIndex), options);
                if (remaining <= 0D) {
                    return column;
                }

                column++;
            }

            return column;
        }

        private static int ResolveLastVisualRow(int startRow, int heightPixels, IReadOnlyDictionary<int, ExcelRowSnapshot> rows, ExcelImageExportOptions options) {
            double remaining = Math.Max(1D, heightPixels);
            int row = startRow;
            while (row < 1048576) {
                rows.TryGetValue(row, out ExcelRowSnapshot? definition);
                remaining -= ResolveRowHeight(definition, options);
                if (remaining <= 0D) {
                    return row;
                }

                row++;
            }

            return row;
        }

        private static double ResolveColumnWidth(ExcelColumnSnapshot? definition, ExcelImageExportOptions options) {
            if (definition?.Width == null) {
                return options.DefaultColumnWidthPixels;
            }

            return Math.Max(1D, Math.Round((definition.Width.Value * 7D) + 5D, 2));
        }

        private static double ResolveRowHeight(ExcelRowSnapshot? definition, ExcelImageExportOptions options) {
            if (definition?.Height == null) {
                return options.DefaultRowHeightPixels;
            }

            return Math.Max(1D, Math.Round(definition.Height.Value * 96D / 72D, 2));
        }

        private static void WriteImageFile(string path, byte[] bytes) {
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("Output path cannot be null or whitespace.", nameof(path));
            }

            string fullPath = Path.GetFullPath(path);
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrWhiteSpace(directory)) {
                Directory.CreateDirectory(directory!);
            }

            File.WriteAllBytes(fullPath, bytes);
        }

        private static void WriteImageStream(Stream stream, byte[] bytes) {
            if (stream == null) {
                throw new ArgumentNullException(nameof(stream));
            }

            stream.Write(bytes, 0, bytes.Length);
        }

        private sealed class WorksheetImageRangeResolution {
            internal WorksheetImageRangeResolution(string range, IReadOnlyList<OfficeImageExportDiagnostic> diagnostics) {
                Range = range;
                Diagnostics = diagnostics;
            }

            internal string Range { get; }
            internal IReadOnlyList<OfficeImageExportDiagnostic> Diagnostics { get; }
        }
    }
}
