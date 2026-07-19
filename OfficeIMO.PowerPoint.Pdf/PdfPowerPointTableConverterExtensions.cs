using A = DocumentFormat.OpenXml.Drawing;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

/// <summary>
/// Converts structured logical PDF tables into PowerPoint tables.
/// </summary>
public static partial class PowerPointPdfConverterExtensions {
    /// <summary>Imports logical PDF tables into a new PowerPoint presentation at <paramref name="presentationPath"/>.</summary>
    public static PdfPowerPointTableImportReport SaveTablesAsPowerPoint(
        this PdfCore.PdfLogicalDocument document,
        string presentationPath,
        PdfPowerPointTableImportOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(presentationPath)) throw new ArgumentException("Presentation path cannot be empty.", nameof(presentationPath));

        PdfPowerPointTableImportResult result = document.ImportTablesToPowerPointPresentationResult(options);
        using (result.Value) {
            result.Value.Save(presentationPath);
        }
        return result.Report;
    }

    /// <summary>Imports logical PDF tables into a PowerPoint presentation written to a caller-owned stream.</summary>
    public static PdfPowerPointTableImportReport SaveTablesAsPowerPoint(
        this PdfCore.PdfLogicalDocument document,
        Stream presentationStream,
        PdfPowerPointTableImportOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (presentationStream == null) throw new ArgumentNullException(nameof(presentationStream));
        if (!presentationStream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(presentationStream));

        PdfPowerPointTableImportResult result = document.ImportTablesToPowerPointPresentationResult(options);
        using (result.Value) {
            result.Value.Save(presentationStream);
        }
        return result.Report;
    }

    /// <summary>Imports logical PDF tables into a new editable PowerPoint presentation.</summary>
    public static PptCore.PowerPointPresentation ImportTablesToPowerPointPresentation(
        this PdfCore.PdfLogicalDocument document,
        PdfPowerPointTableImportOptions? options = null) => document.ImportTablesToPowerPointPresentationResult(options).Value;

    /// <summary>Imports logical PDF tables into an editable PowerPoint presentation plus an explicit table-scope report.</summary>
    public static PdfPowerPointTableImportResult ImportTablesToPowerPointPresentationResult(
        this PdfCore.PdfLogicalDocument document,
        PdfPowerPointTableImportOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        PptCore.PowerPointPresentation presentation = PptCore.PowerPointPresentation.Create();
        IReadOnlyList<PdfPowerPointTableImportEntry> entries = ImportTables(document, presentation, options ?? new PdfPowerPointTableImportOptions());
        PdfCore.PdfTableExtractionScopeReport sourceScope = PdfCore.PdfLogicalTableAnalysis.AnalyzeExtractionScope(document);
        return new PdfPowerPointTableImportResult(presentation, new PdfPowerPointTableImportReport(entries, sourceScope));
    }

    /// <summary>Asynchronously imports logical PDF tables into a PowerPoint presentation written to a file.</summary>
    public static async Task<PdfPowerPointTableImportReport> SaveTablesAsPowerPointAsync(
        this PdfCore.PdfLogicalDocument document,
        string presentationPath,
        PdfPowerPointTableImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(presentationPath)) throw new ArgumentException("Presentation path cannot be empty.", nameof(presentationPath));
        cancellationToken.ThrowIfCancellationRequested();
        PdfPowerPointTableImportResult result = document.ImportTablesToPowerPointPresentationResult(options);
        using (result.Value) {
            await result.Value.SaveAsync(presentationPath, cancellationToken).ConfigureAwait(false);
        }
        return result.Report;
    }

    /// <summary>Asynchronously imports logical PDF tables into a PowerPoint presentation written to a caller-owned stream.</summary>
    public static async Task<PdfPowerPointTableImportReport> SaveTablesAsPowerPointAsync(
        this PdfCore.PdfLogicalDocument document,
        Stream presentationStream,
        PdfPowerPointTableImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (presentationStream == null) throw new ArgumentNullException(nameof(presentationStream));
        if (!presentationStream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(presentationStream));
        cancellationToken.ThrowIfCancellationRequested();
        PdfPowerPointTableImportResult result = document.ImportTablesToPowerPointPresentationResult(options);
        using (result.Value) {
            await result.Value.SaveAsync(presentationStream, cancellationToken).ConfigureAwait(false);
        }
        return result.Report;
    }

    private static IReadOnlyList<PdfPowerPointTableImportEntry> ImportTables(
        PdfCore.PdfLogicalDocument document,
        PptCore.PowerPointPresentation presentation,
        PdfPowerPointTableImportOptions options) {
        IReadOnlyList<PdfCore.PdfLogicalTableExtraction> tables = PdfCore.PdfLogicalTableAnalysis.ExtractTables(document, options.MaxRows);
        if (tables.Count == 0) {
            AddEmptyPresentationSlide(presentation, options);
            return Array.Empty<PdfPowerPointTableImportEntry>();
        }

        var results = new List<PdfPowerPointTableImportEntry>(tables.Count);
        for (int i = 0; i < tables.Count; i++) {
            PdfCore.PdfLogicalTableExtraction extraction = tables[i];
            PdfCore.PdfLogicalTableData data = extraction.Data;
            bool headerRowIncluded = options.IncludeColumnHeaderRows && HasHeaderRow(data);
            List<TableSegment> segments = BuildTableSegments(data, options);
            for (int segmentIndex = 0; segmentIndex < segments.Count; segmentIndex++) {
                TableSegment segment = segments[segmentIndex];
                int tableRowCount = segment.RowCount + (headerRowIncluded ? 1 : 0);
                if (tableRowCount <= 0) {
                    continue;
                }

                int slideIndex = presentation.Slides.Count == 1 && results.Count == 0 ? 0 : presentation.Slides.Count;
                PptCore.PowerPointSlide slide = presentation.AddSlide();

                if (options.IncludeSourceTitles) {
                    slide.AddTitle(BuildTitle(extraction, segmentIndex, segments.Count));
                }

                PptCore.PowerPointTable table = slide.AddTable(
                    tableRowCount,
                    segment.ColumnCount,
                    options.TableStyle,
                    options.TableLeft,
                    options.TableTop,
                    options.TableWidth,
                    options.TableHeight);
                PopulateTable(table, extraction.Table, data, segment, headerRowIncluded, options);

                results.Add(new PdfPowerPointTableImportEntry(
                    extraction.PageIndex,
                    extraction.PageNumber,
                    extraction.TableIndex,
                    extraction.DetectionKind,
                    slideIndex,
                    segmentIndex,
                    segments.Count,
                    segment.RowStartIndex,
                    segment.ColumnStartIndex,
                    data.Columns.Count,
                    segment.ColumnCount,
                    segment.RowCount,
                    data.TotalRowCount,
                    data.Truncated,
                    headerRowIncluded));
            }
        }

        if (results.Count == 0) {
            AddEmptyPresentationSlide(presentation, options);
        }

        return results.AsReadOnly();
    }

    private static void AddEmptyPresentationSlide(PptCore.PowerPointPresentation presentation, PdfPowerPointTableImportOptions options) {
        PptCore.PowerPointSlide slide = presentation.AddSlide();
        string title = string.IsNullOrWhiteSpace(options.EmptyPresentationTitle)
            ? "PDF Tables"
            : options.EmptyPresentationTitle;
        string message = string.IsNullOrWhiteSpace(options.EmptyPresentationMessage)
            ? "No PDF tables detected."
            : options.EmptyPresentationMessage;

        slide.AddTitle(title);
        slide.AddTextBox(message);
    }

    private static bool HasHeaderRow(PdfCore.PdfLogicalTableData data) {
        return data.Columns.Count > 0
            && (data.Structure.HasHeaderRow || data.Structure.IsKeyValueTable)
            && data.Columns.Any(column => !string.IsNullOrWhiteSpace(column));
    }

    private static string BuildTitle(PdfCore.PdfLogicalTableExtraction extraction, int segmentIndex, int segmentCount) {
        string title = "PDF page "
            + extraction.PageNumber.ToString(CultureInfo.InvariantCulture)
            + ", table "
            + (extraction.TableIndex + 1).ToString(CultureInfo.InvariantCulture);
        return segmentCount > 1
            ? title + " (part " + (segmentIndex + 1).ToString(CultureInfo.InvariantCulture) + " of " + segmentCount.ToString(CultureInfo.InvariantCulture) + ")"
            : title;
    }

    private static List<TableSegment> BuildTableSegments(PdfCore.PdfLogicalTableData data, PdfPowerPointTableImportOptions options) {
        int sourceColumnCount = Math.Max(data.Columns.Count, 1);
        int columnLimit = options.MaxColumnsPerSlide > 0
            ? Math.Min(options.MaxColumnsPerSlide, sourceColumnCount)
            : sourceColumnCount;
        int rowLimit = options.MaxRowsPerSlide > 0
            ? Math.Min(options.MaxRowsPerSlide, Math.Max(data.Rows.Count, 1))
            : Math.Max(data.Rows.Count, 1);

        var columnSegments = new List<TableRange>();
        for (int columnStart = 0; columnStart < sourceColumnCount; columnStart += columnLimit) {
            columnSegments.Add(new TableRange(columnStart, Math.Min(columnLimit, sourceColumnCount - columnStart)));
        }

        var rowSegments = new List<TableRange>();
        if (data.Rows.Count == 0) {
            rowSegments.Add(new TableRange(0, 0));
        } else {
            for (int rowStart = 0; rowStart < data.Rows.Count; rowStart += rowLimit) {
                rowSegments.Add(new TableRange(rowStart, Math.Min(rowLimit, data.Rows.Count - rowStart)));
            }
        }

        var segments = new List<TableSegment>(columnSegments.Count * rowSegments.Count);
        for (int rowIndex = 0; rowIndex < rowSegments.Count; rowIndex++) {
            TableRange row = rowSegments[rowIndex];
            for (int columnIndex = 0; columnIndex < columnSegments.Count; columnIndex++) {
                TableRange column = columnSegments[columnIndex];
                segments.Add(new TableSegment(row.StartIndex, row.Count, column.StartIndex, column.Count));
            }
        }

        return segments;
    }

    private static void PopulateTable(
        PptCore.PowerPointTable table,
        PdfCore.PdfLogicalTable sourceTable,
        PdfCore.PdfLogicalTableData data,
        TableSegment segment,
        bool headerRowIncluded,
        PdfPowerPointTableImportOptions options) {
        table.HeaderRow = headerRowIncluded;
        table.BandedRows = options.BandedRows;

        int rowOffset = headerRowIncluded ? 1 : 0;
        if (headerRowIncluded) {
            WriteRow(table, 0, data.Columns, segment.ColumnStartIndex, data, alignNumericColumns: false);
        }

        for (int rowIndex = 0; rowIndex < segment.RowCount; rowIndex++) {
            WriteRow(
                table,
                rowIndex + rowOffset,
                data.Rows[segment.RowStartIndex + rowIndex],
                segment.ColumnStartIndex,
                data,
                options.AlignNumericColumns);
        }

        ApplyTableSizing(table, sourceTable, segment);
    }

    private static void ApplyTableSizing(
        PptCore.PowerPointTable table,
        PdfCore.PdfLogicalTable sourceTable,
        TableSegment segment) {
        if (TryGetColumnWidthRatios(sourceTable, segment, out double[] ratios)) {
            table.SetColumnWidthsByRatio(ratios);
        } else {
            table.SetColumnWidthsEvenly();
        }

        table.SetRowHeightsEvenly();
    }

    private static bool TryGetColumnWidthRatios(
        PdfCore.PdfLogicalTable sourceTable,
        TableSegment segment,
        out double[] ratios) {
        ratios = Array.Empty<double>();
        if (segment.ColumnCount <= 0 ||
            sourceTable.Columns.Count < segment.ColumnStartIndex + segment.ColumnCount) {
            return false;
        }

        var values = new double[segment.ColumnCount];
        for (int columnIndex = 0; columnIndex < segment.ColumnCount; columnIndex++) {
            PdfCore.PdfLogicalTableColumn sourceColumn = sourceTable.Columns[segment.ColumnStartIndex + columnIndex];
            double width = sourceColumn.To - sourceColumn.From;
            if (double.IsNaN(width) || double.IsInfinity(width) || width <= 0) {
                return false;
            }

            values[columnIndex] = width;
        }

        ratios = values;
        return true;
    }

    private static void WriteRow(
        PptCore.PowerPointTable table,
        int rowIndex,
        IReadOnlyList<string> values,
        int sourceColumnStartIndex,
        PdfCore.PdfLogicalTableData data,
        bool alignNumericColumns) {
        for (int columnIndex = 0; columnIndex < table.Columns; columnIndex++) {
            int sourceColumnIndex = sourceColumnStartIndex + columnIndex;
            string value = sourceColumnIndex < values.Count ? values[sourceColumnIndex] : string.Empty;
            PptCore.PowerPointTableCell cell = table.GetCell(rowIndex, columnIndex);
            cell.Text = value ?? string.Empty;
            if (alignNumericColumns && data.IsNumericColumn(sourceColumnIndex)) {
                cell.HorizontalAlignment = A.TextAlignmentTypeValues.Right;
            }
        }
    }

    private readonly struct TableRange {
        public TableRange(int startIndex, int count) {
            StartIndex = startIndex;
            Count = count;
        }

        public int StartIndex { get; }

        public int Count { get; }
    }

    private readonly struct TableSegment {
        public TableSegment(int rowStartIndex, int rowCount, int columnStartIndex, int columnCount) {
            RowStartIndex = rowStartIndex;
            RowCount = rowCount;
            ColumnStartIndex = columnStartIndex;
            ColumnCount = columnCount;
        }

        public int RowStartIndex { get; }

        public int RowCount { get; }

        public int ColumnStartIndex { get; }

        public int ColumnCount { get; }
    }
}
