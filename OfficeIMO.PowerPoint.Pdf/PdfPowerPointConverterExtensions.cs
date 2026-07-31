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
    /// <summary>Converts an opened PDF into a PowerPoint presentation.</summary>
    public static PptCore.PowerPointPresentation ToPowerPointPresentation(
        this PdfCore.PdfDocument document,
        PdfPowerPointImportOptions? options = null) =>
        document.ToPowerPointPresentationResult(options).Value;

    /// <summary>Converts an opened PDF into a PowerPoint presentation with conversion diagnostics.</summary>
    public static PdfPowerPointConversionResult ToPowerPointPresentationResult(
        this PdfCore.PdfDocument document,
        PdfPowerPointImportOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        PdfPowerPointImportOptions operation = options ?? new PdfPowerPointImportOptions();
        if (operation.Mode == PdfPowerPointImportMode.EditableTables) {
            PdfCore.PdfLogicalDocument logical = operation.PageSelection == null
                ? document.Read.Logical()
                : document.Read.Logical(operation.PageSelection);
            return logical.ToPowerPointPresentationResult(operation);
        }

        if (operation.Mode == PdfPowerPointImportMode.HybridVisualAndEditableTables) {
            return ImportHybridPages(document, operation);
        }

        return ImportVisualPages(document, operation);
    }

    /// <summary>Converts an opened PDF and saves the PowerPoint presentation to a file.</summary>
    public static PdfPowerPointConversionReport SaveAsPowerPoint(
        this PdfCore.PdfDocument document,
        string presentationPath,
        PdfPowerPointImportOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(presentationPath)) throw new ArgumentException("Presentation path cannot be empty.", nameof(presentationPath));
        PdfPowerPointConversionResult result = document.ToPowerPointPresentationResult(options);
        using (result.Value) result.Value.Save(presentationPath);
        return result.Report;
    }

    /// <summary>Converts an opened PDF and saves the PowerPoint presentation to a caller-owned stream.</summary>
    public static PdfPowerPointConversionReport SaveAsPowerPoint(
        this PdfCore.PdfDocument document,
        Stream presentationStream,
        PdfPowerPointImportOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (presentationStream == null) throw new ArgumentNullException(nameof(presentationStream));
        if (!presentationStream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(presentationStream));
        PdfPowerPointConversionResult result = document.ToPowerPointPresentationResult(options);
        using (result.Value) result.Value.Save(presentationStream);
        return result.Report;
    }

    /// <summary>Converts an opened PDF and asynchronously saves the PowerPoint presentation to a file.</summary>
    public static async Task<PdfPowerPointConversionReport> SaveAsPowerPointAsync(
        this PdfCore.PdfDocument document,
        string presentationPath,
        PdfPowerPointImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(presentationPath)) throw new ArgumentException("Presentation path cannot be empty.", nameof(presentationPath));
        cancellationToken.ThrowIfCancellationRequested();
        PdfPowerPointConversionResult result = document.ToPowerPointPresentationResult(options);
        using (result.Value) await result.Value.SaveAsync(presentationPath, cancellationToken).ConfigureAwait(false);
        return result.Report;
    }

    /// <summary>Converts an opened PDF and asynchronously saves the PowerPoint presentation to a caller-owned stream.</summary>
    public static async Task<PdfPowerPointConversionReport> SaveAsPowerPointAsync(
        this PdfCore.PdfDocument document,
        Stream presentationStream,
        PdfPowerPointImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (presentationStream == null) throw new ArgumentNullException(nameof(presentationStream));
        if (!presentationStream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(presentationStream));
        cancellationToken.ThrowIfCancellationRequested();
        PdfPowerPointConversionResult result = document.ToPowerPointPresentationResult(options);
        using (result.Value) await result.Value.SaveAsync(presentationStream, cancellationToken).ConfigureAwait(false);
        return result.Report;
    }

    /// <summary>Imports logical PDF tables into a new PowerPoint presentation at <paramref name="presentationPath"/>.</summary>
    public static PdfPowerPointConversionReport SaveAsPowerPoint(
        this PdfCore.PdfLogicalDocument document,
        string presentationPath,
        PdfPowerPointImportOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(presentationPath)) throw new ArgumentException("Presentation path cannot be empty.", nameof(presentationPath));

        PdfPowerPointConversionResult result = document.ToPowerPointPresentationResult(options);
        using (result.Value) {
            result.Value.Save(presentationPath);
        }
        return result.Report;
    }

    /// <summary>Imports logical PDF tables into a PowerPoint presentation written to a caller-owned stream.</summary>
    public static PdfPowerPointConversionReport SaveAsPowerPoint(
        this PdfCore.PdfLogicalDocument document,
        Stream presentationStream,
        PdfPowerPointImportOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (presentationStream == null) throw new ArgumentNullException(nameof(presentationStream));
        if (!presentationStream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(presentationStream));

        PdfPowerPointConversionResult result = document.ToPowerPointPresentationResult(options);
        using (result.Value) {
            result.Value.Save(presentationStream);
        }
        return result.Report;
    }

    /// <summary>Imports logical PDF tables into a new editable PowerPoint presentation.</summary>
    public static PptCore.PowerPointPresentation ToPowerPointPresentation(
        this PdfCore.PdfLogicalDocument document,
        PdfPowerPointImportOptions? options = null) => document.ToPowerPointPresentationResult(options).Value;

    /// <summary>Imports logical PDF tables into an editable PowerPoint presentation plus an explicit table-scope report.</summary>
    public static PdfPowerPointConversionResult ToPowerPointPresentationResult(
        this PdfCore.PdfLogicalDocument document,
        PdfPowerPointImportOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        PdfPowerPointImportOptions operation = options ?? PdfPowerPointImportOptions.CreateEditableTables();
        if (operation.Mode != PdfPowerPointImportMode.EditableTables) {
            throw new InvalidOperationException("Visual PDF page import requires the opened PdfDocument so page rendering can use the original PDF bytes.");
        }
        PptCore.PowerPointPresentation presentation = PptCore.PowerPointPresentation.Create();
        IReadOnlyList<PdfPowerPointTableImportEntry> entries = ImportTables(document, presentation, operation);
        PdfCore.PdfTableExtractionScopeReport sourceScope = PdfCore.PdfLogicalTableAnalysis.AnalyzeExtractionScope(document);
        return new PdfPowerPointConversionResult(presentation, new PdfPowerPointConversionReport(entries, sourceScope));
    }

    /// <summary>Asynchronously imports logical PDF tables into a PowerPoint presentation written to a file.</summary>
    public static async Task<PdfPowerPointConversionReport> SaveAsPowerPointAsync(
        this PdfCore.PdfLogicalDocument document,
        string presentationPath,
        PdfPowerPointImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(presentationPath)) throw new ArgumentException("Presentation path cannot be empty.", nameof(presentationPath));
        cancellationToken.ThrowIfCancellationRequested();
        PdfPowerPointConversionResult result = document.ToPowerPointPresentationResult(options);
        using (result.Value) {
            await result.Value.SaveAsync(presentationPath, cancellationToken).ConfigureAwait(false);
        }
        return result.Report;
    }

    /// <summary>Asynchronously imports logical PDF tables into a PowerPoint presentation written to a caller-owned stream.</summary>
    public static async Task<PdfPowerPointConversionReport> SaveAsPowerPointAsync(
        this PdfCore.PdfLogicalDocument document,
        Stream presentationStream,
        PdfPowerPointImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (presentationStream == null) throw new ArgumentNullException(nameof(presentationStream));
        if (!presentationStream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(presentationStream));
        cancellationToken.ThrowIfCancellationRequested();
        PdfPowerPointConversionResult result = document.ToPowerPointPresentationResult(options);
        using (result.Value) {
            await result.Value.SaveAsync(presentationStream, cancellationToken).ConfigureAwait(false);
        }
        return result.Report;
    }

    private static IReadOnlyList<PdfPowerPointTableImportEntry> ImportTables(
        PdfCore.PdfLogicalDocument document,
        PptCore.PowerPointPresentation presentation,
        PdfPowerPointImportOptions options) {
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

    private static PdfPowerPointConversionResult ImportVisualPages(
        PdfCore.PdfDocument document,
        PdfPowerPointImportOptions options) {
        IReadOnlyList<PdfCore.PdfPageRenderResult> rendered = RenderPages(document, options);
        PptCore.PowerPointPresentation presentation = PptCore.PowerPointPresentation.Create();
        ConfigureSlideSize(presentation, rendered);

        var entries = new List<PdfPowerPointVisualPageEntry>(rendered.Count);
        for (int i = 0; i < rendered.Count; i++) {
            PptCore.PowerPointSlide slide = presentation.AddSlide();
            AddVisualPage(presentation, slide, rendered[i]);
            entries.Add(new PdfPowerPointVisualPageEntry(rendered[i], i));
        }

        if (rendered.Count == 0) {
            PptCore.PowerPointSlide slide = presentation.AddSlide();
            slide.AddTitle("PDF");
            slide.AddTextBox("No PDF pages were selected.");
        }

        return new PdfPowerPointConversionResult(presentation, new PdfPowerPointConversionReport(entries));
    }

    private static PdfPowerPointConversionResult ImportHybridPages(
        PdfCore.PdfDocument document,
        PdfPowerPointImportOptions options) {
        IReadOnlyList<PdfCore.PdfPageRenderResult> rendered = RenderPages(document, options);
        PdfCore.PdfLogicalDocument logical = options.PageSelection == null
            ? document.Read.Logical()
            : document.Read.Logical(options.PageSelection);
        PptCore.PowerPointPresentation presentation = PptCore.PowerPointPresentation.Create();
        ConfigureSlideSize(presentation, rendered);
        var visualEntries = new List<PdfPowerPointVisualPageEntry>(rendered.Count);
        var tableEntries = new List<PdfPowerPointTableImportEntry>();

        for (int pageIndex = 0; pageIndex < rendered.Count; pageIndex++) {
            PdfCore.PdfPageRenderResult render = rendered[pageIndex];
            PdfCore.PdfLogicalPage? page = pageIndex < logical.Pages.Count ? logical.Pages[pageIndex] : null;
            IReadOnlyList<PdfCore.PdfLogicalTableExtraction> tables = page == null
                ? Array.Empty<PdfCore.PdfLogicalTableExtraction>()
                : PdfCore.PdfLogicalTableAnalysis.ExtractTables(page, options.MaxRows);
            bool addedTableSlide = false;
            for (int tableIndex = 0; tableIndex < tables.Count; tableIndex++) {
                PdfCore.PdfLogicalTableExtraction extraction = tables[tableIndex];
                PdfCore.PdfLogicalTableData data = extraction.Data;
                if (data.Columns.Count <= 0) continue;
                List<TableSegment> segments = BuildTableSegments(data, options);
                for (int segmentIndex = 0; segmentIndex < segments.Count; segmentIndex++) {
                    TableSegment segment = segments[segmentIndex];
                    bool headerRowIncluded = options.IncludeColumnHeaderRows
                        && HasHeaderRow(data)
                        && segment.RowStartIndex == 0;
                    int rowCount = segment.RowCount + (headerRowIncluded ? 1 : 0);
                    if (rowCount <= 0 || segment.ColumnCount <= 0) continue;

                    int slideIndex = presentation.Slides.Count;
                    PptCore.PowerPointSlide slide = presentation.AddSlide();
                    AddVisualPage(presentation, slide, render);
                    visualEntries.Add(new PdfPowerPointVisualPageEntry(render, slideIndex));
                    (long left, long top, long width, long height) = GetHybridTableBounds(
                        render,
                        page!,
                        extraction.Table,
                        data,
                        segment,
                        headerRowIncluded,
                        presentation.SlideSize.WidthPoints,
                        presentation.SlideSize.HeightPoints);
                    PptCore.PowerPointTable table = slide.AddTable(
                        rowCount,
                        segment.ColumnCount,
                        options.TableStyle,
                        left,
                        top,
                        width,
                        height);
                    PopulateTable(table, extraction.Table, data, segment, headerRowIncluded, options);
                    tableEntries.Add(new PdfPowerPointTableImportEntry(
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
                    addedTableSlide = true;
                }
            }

            if (!addedTableSlide) {
                int slideIndex = presentation.Slides.Count;
                PptCore.PowerPointSlide slide = presentation.AddSlide();
                AddVisualPage(presentation, slide, render);
                visualEntries.Add(new PdfPowerPointVisualPageEntry(render, slideIndex));
            }
        }

        if (rendered.Count == 0) {
            PptCore.PowerPointSlide slide = presentation.AddSlide();
            slide.AddTitle("PDF");
            slide.AddTextBox("No PDF pages were selected.");
        }

        PdfCore.PdfTableExtractionScopeReport scope = PdfCore.PdfLogicalTableAnalysis.AnalyzeExtractionScope(logical);
        return new PdfPowerPointConversionResult(
            presentation,
            new PdfPowerPointConversionReport(tableEntries, visualEntries, scope));
    }

    private static IReadOnlyList<PdfCore.PdfPageRenderResult> RenderPages(
        PdfCore.PdfDocument document,
        PdfPowerPointImportOptions options) {
        var renderOptions = new PdfCore.PdfPageRenderOptions {
            Format = PdfCore.PdfPageRenderFormat.Png,
            Dpi = options.Dpi,
            MaxPages = options.MaxPages,
            MaxPixelsPerPage = options.MaxPixelsPerPage,
            MaxOutputBytesPerPage = options.MaxOutputBytesPerPage,
            MaxTotalOutputBytes = options.MaxTotalOutputBytes,
            ContinueOnError = true
        };
        return document.Read.RenderPages(
            options.PageSelection,
            renderOptions);
    }

    private static void ConfigureSlideSize(
        PptCore.PowerPointPresentation presentation,
        IReadOnlyList<PdfCore.PdfPageRenderResult> rendered) {
        PdfCore.PdfPageRenderResult? reference = rendered.FirstOrDefault(static page => page.Succeeded);
        if (reference != null) {
            const double maximumSlideDimensionPoints = 720D;
            double width = reference.Width >= reference.Height
                ? maximumSlideDimensionPoints
                : maximumSlideDimensionPoints * reference.Width / Math.Max(1D, reference.Height);
            double height = reference.Height >= reference.Width
                ? maximumSlideDimensionPoints
                : maximumSlideDimensionPoints * reference.Height / Math.Max(1D, reference.Width);
            presentation.SlideSize.SetSizePoints(width, height);
        }
    }

    private static void AddVisualPage(
        PptCore.PowerPointPresentation presentation,
        PptCore.PowerPointSlide slide,
        PdfCore.PdfPageRenderResult page) {
        byte[]? bytes = page.Bytes;
        if (bytes != null && page.Width > 0 && page.Height > 0) {
            VisualPagePlacement placement = GetVisualPagePlacement(
                page,
                presentation.SlideSize.WidthPoints,
                presentation.SlideSize.HeightPoints);
            using var image = new MemoryStream(bytes, writable: false);
            slide.AddPicturePoints(image, PptCore.ImagePartType.Png, placement.Left, placement.Top, placement.Width, placement.Height);
        } else {
            slide.AddTitle("PDF page " + page.PageNumber.ToString(CultureInfo.InvariantCulture));
            slide.AddTextBox("This page could not be rendered by the managed PDF renderer.");
        }
    }

    private static (long Left, long Top, long Width, long Height) GetHybridTableBounds(
        PdfCore.PdfPageRenderResult render,
        PdfCore.PdfLogicalPage page,
        PdfCore.PdfLogicalTable table,
        PdfCore.PdfLogicalTableData data,
        TableSegment segment,
        bool headerRowIncluded,
        double slideWidthPoints,
        double slideHeightPoints) {
        const double emusPerPoint = 12700D;
        VisualPagePlacement placement = GetVisualPagePlacement(render, slideWidthPoints, slideHeightPoints);
        int firstColumn = Math.Min(segment.ColumnStartIndex, Math.Max(0, table.Columns.Count - 1));
        int lastColumn = Math.Min(table.Columns.Count, segment.ColumnStartIndex + segment.ColumnCount) - 1;
        double leftPoints = table.Columns.Count == 0 ? 0D : table.Columns[firstColumn].From;
        double rightPoints = lastColumn < firstColumn ? page.Width : table.Columns[lastColumn].To;

        int sourceRowCount = Math.Max(1, table.Rows.Count);
        double sourceRowHeight = Math.Max(0D, table.YTop - table.YBottom) / sourceRowCount;
        bool sourceHeaderIncluded = headerRowIncluded && data.Structure.HasHeaderRow;
        int sourceRowStart = sourceHeaderIncluded
            ? 0
            : Math.Min(sourceRowCount - 1, data.Structure.BodyStartRowIndex + segment.RowStartIndex);
        int sourceRows = Math.Max(1, segment.RowCount + (sourceHeaderIncluded ? 1 : 0));
        double sourceTop = table.YTop - sourceRowStart * sourceRowHeight;
        double sourceBottom = Math.Max(table.YBottom, sourceTop - sourceRows * sourceRowHeight);
        double topPoints = Math.Max(0D, page.Height - sourceTop);
        double widthPoints = Math.Max(1D, rightPoints - leftPoints);
        double heightPoints = Math.Max(1D, sourceTop - sourceBottom);
        double xScale = placement.Width / Math.Max(1D, page.Width);
        double yScale = placement.Height / Math.Max(1D, page.Height);
        return (
            (long)Math.Round((placement.Left + leftPoints * xScale) * emusPerPoint),
            (long)Math.Round((placement.Top + topPoints * yScale) * emusPerPoint),
            (long)Math.Round(Math.Min(placement.Width, widthPoints * xScale) * emusPerPoint),
            (long)Math.Round(Math.Min(placement.Height, heightPoints * yScale) * emusPerPoint));
    }

    private static VisualPagePlacement GetVisualPagePlacement(
        PdfCore.PdfPageRenderResult page,
        double slideWidthPoints,
        double slideHeightPoints) {
        double scale = Math.Min(slideWidthPoints / Math.Max(1D, page.Width), slideHeightPoints / Math.Max(1D, page.Height));
        double width = page.Width * scale;
        double height = page.Height * scale;
        return new VisualPagePlacement(
            (slideWidthPoints - width) / 2D,
            (slideHeightPoints - height) / 2D,
            width,
            height);
    }

    private static void AddEmptyPresentationSlide(PptCore.PowerPointPresentation presentation, PdfPowerPointImportOptions options) {
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

    private static List<TableSegment> BuildTableSegments(PdfCore.PdfLogicalTableData data, PdfPowerPointImportOptions options) {
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
        PdfPowerPointImportOptions options) {
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

    private readonly struct VisualPagePlacement {
        internal VisualPagePlacement(double left, double top, double width, double height) {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }

        internal double Left { get; }
        internal double Top { get; }
        internal double Width { get; }
        internal double Height { get; }
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
