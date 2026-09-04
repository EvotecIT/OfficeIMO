using System.Globalization;
using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>
/// Shared analysis helpers for logical PDF tables produced by <see cref="PdfDocumentReadResult"/>.
/// </summary>
public static class PdfLogicalTableAnalysis {
    private const int DefaultMaximumScopeAnalysisComparisons = 10_000;
    private const int MaximumScopeComparisonTextCharacters = 512;
    private const int MaximumScopeSourceCharactersPerValue = 2048;
    /// <summary>
    /// Establishes structured extraction metadata for a logical PDF table from structural evidence.
    /// </summary>
    /// <param name="table">Logical table to inspect.</param>
    /// <returns>Column names, body-row boundary, and table-shape flags for structured consumers.</returns>
    public static PdfLogicalTableStructure Analyze(PdfLogicalTable table) =>
        Analyze(table, CancellationToken.None);

    private static PdfLogicalTableStructure Analyze(
        PdfLogicalTable table,
        CancellationToken cancellationToken) {
        Guard.NotNull(table, nameof(table));
        cancellationToken.ThrowIfCancellationRequested();

        int columnCount = GetColumnCount(table, cancellationToken);
        string[]? headerColumns = DetectHeaderColumns(table, cancellationToken);
        bool hasHeader = headerColumns != null && headerColumns.Length == columnCount;
        int bodyStartRowIndex = hasHeader ? 1 : 0;
        int totalBodyRowCount = Math.Max(0, table.Rows.Count - bodyStartRowIndex);
        IReadOnlyList<string> columns = hasHeader
            ? headerColumns!
            : BuildUnnamedColumns(columnCount);
        PdfLogicalTableSchemaKind schemaKind = hasHeader
            ? PdfLogicalTableSchemaKind.HeaderRow
            : PdfLogicalTableSchemaKind.Unknown;
        IReadOnlyList<PdfInferenceEvidence> schemaEvidence = BuildSchemaEvidence(table, hasHeader);
        double schemaConfidence = hasHeader
            ? table.Evidence.Any(static evidence => evidence.Code == "table.tagged-header-row") ? 0.99D : 0.95D
            : 0D;

        return new PdfLogicalTableStructure(
            columnCount,
            columns,
            bodyStartRowIndex,
            totalBodyRowCount,
            hasHeader,
            schemaKind,
            schemaConfidence,
            schemaEvidence);
    }

    /// <summary>
    /// Extracts a normalized, structured table view for document readers and text emitters.
    /// </summary>
    /// <param name="table">Logical table to inspect.</param>
    /// <param name="maxRows">Maximum number of body rows to return. Values less than or equal to zero return all body rows.</param>
    /// <returns>Inferred columns, normalized body rows, numeric-column flags, and truncation metadata.</returns>
    public static PdfLogicalTableData Extract(PdfLogicalTable table, int maxRows = 0) =>
        Extract(table, maxRows, CancellationToken.None);

    private static PdfLogicalTableData Extract(
        PdfLogicalTable table,
        int maxRows,
        CancellationToken cancellationToken) {
        Guard.NotNull(table, nameof(table));
        cancellationToken.ThrowIfCancellationRequested();

        PdfLogicalTableStructure structure = Analyze(table, cancellationToken);
        IReadOnlyList<IReadOnlyList<string>> rows = GetBodyRows(table, structure, maxRows, cancellationToken);
        IReadOnlyList<bool> numericColumns = DetectNumericColumns(table, structure, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        PdfLogicalTableDiagnostics diagnostics = PdfLogicalTableDiagnostics.Create(table, structure);
        return new PdfLogicalTableData(
            structure,
            diagnostics,
            rows,
            numericColumns,
            rows.Count < structure.TotalBodyRowCount);
    }

    /// <summary>
    /// Extracts normalized tables from every logical page in document order.
    /// </summary>
    /// <param name="document">Logical PDF document to inspect.</param>
    /// <param name="maxRows">Maximum number of body rows per table. Values less than or equal to zero return all body rows.</param>
    /// <returns>Page-aware normalized table extractions.</returns>
    public static IReadOnlyList<PdfLogicalTableExtraction> ExtractTables(PdfDocumentReadResult document, int maxRows = 0) =>
        ExtractTables(document, maxRows, CancellationToken.None);

    internal static IReadOnlyList<PdfLogicalTableExtraction> ExtractTables(
        PdfDocumentReadResult document,
        int maxRows,
        CancellationToken cancellationToken) {
        Guard.NotNull(document, nameof(document));

        return ExtractTables(document.Pages, maxRows, cancellationToken);
    }

    /// <summary>
    /// Extracts normalized tables from the supplied logical pages in their current order.
    /// </summary>
    /// <param name="pages">Logical pages to inspect.</param>
    /// <param name="maxRows">Maximum number of body rows per table. Values less than or equal to zero return all body rows.</param>
    /// <returns>Page-aware normalized table extractions.</returns>
    public static IReadOnlyList<PdfLogicalTableExtraction> ExtractTables(IReadOnlyList<PdfLogicalPage> pages, int maxRows = 0) =>
        ExtractTables(pages, maxRows, CancellationToken.None);

    private static IReadOnlyList<PdfLogicalTableExtraction> ExtractTables(
        IReadOnlyList<PdfLogicalPage> pages,
        int maxRows,
        CancellationToken cancellationToken) {
        Guard.NotNull(pages, nameof(pages));
        cancellationToken.ThrowIfCancellationRequested();

        var extractions = new List<PdfLogicalTableExtraction>();
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfLogicalPage page = pages[pageIndex];
            for (int tableIndex = 0; tableIndex < page.Tables.Count; tableIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                PdfLogicalTable table = page.Tables[tableIndex];
                extractions.Add(new PdfLogicalTableExtraction(
                    pageIndex,
                    page.PageNumber,
                    tableIndex,
                    table,
                    Extract(table, maxRows, cancellationToken)));
            }
        }

        return extractions.Count == 0 ? Array.Empty<PdfLogicalTableExtraction>() : extractions.AsReadOnly();
    }

    /// <summary>
    /// Extracts normalized tables from a single logical page.
    /// </summary>
    /// <param name="page">Logical page to inspect.</param>
    /// <param name="maxRows">Maximum number of body rows per table. Values less than or equal to zero return all body rows.</param>
    /// <returns>Normalized table extractions for the page.</returns>
    public static IReadOnlyList<PdfLogicalTableExtraction> ExtractTables(PdfLogicalPage page, int maxRows = 0) {
        Guard.NotNull(page, nameof(page));

        return ExtractTables(new[] { page }, maxRows, CancellationToken.None);
    }

    /// <summary>
    /// Describes the source-page content considered by table-only adapters.
    /// </summary>
    /// <param name="document">Logical PDF document to inspect.</param>
    /// <returns>
    /// Table counts plus visible and interactive page content that a table-only adapter will not import.
    /// </returns>
    public static PdfTableExtractionScopeReport AnalyzeExtractionScope(PdfDocumentReadResult document) {
        return AnalyzeExtractionScope(document, DefaultMaximumScopeAnalysisComparisons);
    }

    /// <summary>
    /// Describes table extraction scope while bounding attacker-controlled text/table comparisons.
    /// </summary>
    public static PdfTableExtractionScopeReport AnalyzeExtractionScope(
        PdfDocumentReadResult document,
        int maximumComparisons) {
        Guard.NotNull(document, nameof(document));
        return AnalyzeExtractionScope(
            document.Pages,
            document.OptionalContentGroupCount,
            maximumComparisons);
    }

    /// <summary>
    /// Describes table extraction scope for a selected collection of logical pages.
    /// Document-level optional-content groups are not attributed to individual pages.
    /// </summary>
    /// <param name="pages">Logical pages to inspect.</param>
    /// <returns>Page-scoped table counts plus visible and interactive content outside detected tables.</returns>
    public static PdfTableExtractionScopeReport AnalyzeExtractionScope(IReadOnlyList<PdfLogicalPage> pages) {
        return AnalyzeExtractionScope(pages, DefaultMaximumScopeAnalysisComparisons);
    }

    /// <summary>
    /// Describes table extraction scope for selected logical pages while bounding attacker-controlled text/table comparisons.
    /// Document-level optional-content groups are not attributed to individual pages.
    /// </summary>
    public static PdfTableExtractionScopeReport AnalyzeExtractionScope(
        IReadOnlyList<PdfLogicalPage> pages,
        int maximumComparisons) {
        Guard.NotNull(pages, nameof(pages));
        return AnalyzeExtractionScope(pages, optionalContentGroupCount: 0, maximumComparisons);
    }

    private static PdfTableExtractionScopeReport AnalyzeExtractionScope(
        IReadOnlyList<PdfLogicalPage> pages,
        int optionalContentGroupCount,
        int maximumComparisons) {
#pragma warning disable CA1512 // ThrowIfNegative is unavailable on netstandard2.0 and net472.
        if (maximumComparisons < 0) throw new ArgumentOutOfRangeException(nameof(maximumComparisons));
#pragma warning restore CA1512

        int pagesWithTables = 0;
        int detectedTableCount = 0;
        int nonTableTextBlockCount = 0;
        int vectorPrimitiveCount = 0;
        int imageCount = 0;
        int linkCount = 0;
        int formWidgetCount = 0;
        int annotationCount = 0;
        int pageActionCount = 0;
        int interactiveMediaAnnotationCount = 0;
        int remainingComparisons = Math.Min(maximumComparisons, DefaultMaximumScopeAnalysisComparisons);
        bool analysisTruncated = false;
        var normalizedRows = new Dictionary<PdfLogicalTable, Dictionary<int, ScopeComparisonText>>();

        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            PdfLogicalPage page = pages[pageIndex];
            if (page.Tables.Count > 0) {
                pagesWithTables++;
                detectedTableCount += page.Tables.Count;
            }

            for (int blockIndex = 0; blockIndex < page.TextBlocks.Count; blockIndex++) {
                PdfLogicalTextBlock block = page.TextBlocks[blockIndex];
                ScopeRepresentation representation = IsTextBlockRepresentedByAnyTable(
                    block,
                    page.Tables,
                    normalizedRows,
                    ref remainingComparisons);
                if (representation == ScopeRepresentation.NotRepresented) {
                    nonTableTextBlockCount++;
                } else if (representation == ScopeRepresentation.Incomplete) {
                    analysisTruncated = true;
                }
            }

            vectorPrimitiveCount += page.VectorPrimitiveCount;
            imageCount += page.Images.Count;
            linkCount += page.Links.Count;
            formWidgetCount += page.FormWidgets.Count;
            annotationCount += page.Annotations.Count;
            pageActionCount += page.PageActions.Count;
            interactiveMediaAnnotationCount += page.Annotations.Count(static annotation =>
                IsInteractiveMediaAnnotationSubtype(annotation.Subtype));
        }

        return new PdfTableExtractionScopeReport(
            pages.Count,
            pagesWithTables,
            detectedTableCount,
            nonTableTextBlockCount,
            vectorPrimitiveCount,
            imageCount,
            linkCount,
            formWidgetCount,
            annotationCount,
            pageActionCount,
            optionalContentGroupCount,
            interactiveMediaAnnotationCount,
            analysisTruncated);
    }

    private static bool IsInteractiveMediaAnnotationSubtype(string subtype) =>
        string.Equals(subtype, "Movie", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(subtype, "Sound", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(subtype, "Screen", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(subtype, "RichMedia", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(subtype, "3D", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Detects a header row only when the PDF supplies explicit tagged or typographic structure.
    /// </summary>
    /// <param name="table">Logical table to inspect.</param>
    /// <returns>Header column names when structural evidence establishes the first row as a header; otherwise null.</returns>
    public static IReadOnlyList<string>? DetectHeaderColumns(PdfLogicalTable table) =>
        DetectHeaderColumns(table, CancellationToken.None);

    private static string[]? DetectHeaderColumns(
        PdfLogicalTable table,
        CancellationToken cancellationToken) {
        Guard.NotNull(table, nameof(table));
        cancellationToken.ThrowIfCancellationRequested();

        int columnCount = GetColumnCount(table, cancellationToken);
        if (table.Rows.Count == 0 || columnCount == 0) {
            return null;
        }

        bool hasExplicitHeaderEvidence = table.Evidence.Any(static evidence =>
            string.Equals(evidence.Code, "table.header-emphasis", StringComparison.Ordinal) ||
            string.Equals(evidence.Code, "table.tagged-header-row", StringComparison.Ordinal));
        if (!hasExplicitHeaderEvidence) {
            return null;
        }

        IReadOnlyList<string> firstRow = table.Rows[0];
        var headers = new string[columnCount];
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            headers[columnIndex] = columnIndex < firstRow.Count
                ? firstRow[columnIndex].Trim()
                : string.Empty;
        }

        return headers;
    }

    private static ScopeRepresentation IsTextBlockRepresentedByAnyTable(
        PdfLogicalTextBlock block,
        IReadOnlyList<PdfLogicalTable> tables,
        IDictionary<PdfLogicalTable, Dictionary<int, ScopeComparisonText>> normalizedRows,
        ref int remainingComparisons) {
        ScopeComparisonText? normalizedBlock = null;
        for (int tableIndex = 0; tableIndex < tables.Count; tableIndex++) {
            if (remainingComparisons-- <= 0) return ScopeRepresentation.Incomplete;
            PdfLogicalTable table = tables[tableIndex];
            double top = Math.Max(table.YTop, table.YBottom);
            double bottom = Math.Min(table.YTop, table.YBottom);
            if (block.BaselineY > top + 1D || block.BaselineY < bottom - 1D) {
                continue;
            }

            normalizedBlock ??= NormalizeForScopeComparison(block.Text);
            if (normalizedBlock.Value.Truncated) return ScopeRepresentation.Incomplete;
            string blockText = normalizedBlock.Value.Value;
            if (blockText.Length == 0) {
                return ScopeRepresentation.Represented;
            }

            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                if (remainingComparisons-- <= 0) return ScopeRepresentation.Incomplete;
                ScopeComparisonText normalizedRow = GetNormalizedScopeRow(table, rowIndex, normalizedRows);
                if (normalizedRow.Truncated) return ScopeRepresentation.Incomplete;
                string rowText = normalizedRow.Value;
                if (rowText.Length > 0 &&
                    (ContainsOrdinal(rowText, blockText) ||
                     ContainsOrdinal(blockText, rowText))) {
                    return ScopeRepresentation.Represented;
                }
            }
        }

        return ScopeRepresentation.NotRepresented;
    }

    private static ScopeComparisonText GetNormalizedScopeRow(
        PdfLogicalTable table,
        int rowIndex,
        IDictionary<PdfLogicalTable, Dictionary<int, ScopeComparisonText>> normalizedRows) {
        if (!normalizedRows.TryGetValue(table, out Dictionary<int, ScopeComparisonText>? rows)) {
            rows = new Dictionary<int, ScopeComparisonText>();
            normalizedRows.Add(table, rows);
        }
        if (rows.TryGetValue(rowIndex, out ScopeComparisonText cached)) return cached;

        ScopeComparisonText normalized = NormalizeForScopeComparison(table.Rows[rowIndex]);
        rows.Add(rowIndex, normalized);
        return normalized;
    }

    private static ScopeComparisonText NormalizeForScopeComparison(string? value) {
        if (string.IsNullOrEmpty(value)) return new ScopeComparisonText(string.Empty, truncated: false);

        string normalizedValue = value!;
        var builder = new System.Text.StringBuilder(Math.Min(normalizedValue.Length, MaximumScopeComparisonTextCharacters));
        int inspected = 0;
        for (int index = 0; index < normalizedValue.Length; index++) {
            if (inspected++ == MaximumScopeSourceCharactersPerValue) {
                return new ScopeComparisonText(builder.ToString(), truncated: true);
            }
            char character = normalizedValue[index];
            if (!char.IsWhiteSpace(character)) {
                if (builder.Length == MaximumScopeComparisonTextCharacters) {
                    return new ScopeComparisonText(builder.ToString(), truncated: true);
                }
                builder.Append(char.ToUpperInvariant(character));
            }
        }

        return new ScopeComparisonText(builder.ToString(), truncated: false);
    }

    private static ScopeComparisonText NormalizeForScopeComparison(IReadOnlyList<string> row) {
        var builder = new System.Text.StringBuilder(MaximumScopeComparisonTextCharacters);
        int inspected = 0;
        for (int cellIndex = 0; cellIndex < row.Count; cellIndex++) {
            string value = row[cellIndex] ?? string.Empty;
            for (int index = 0; index < value.Length; index++) {
                if (inspected++ == MaximumScopeSourceCharactersPerValue) {
                    return new ScopeComparisonText(builder.ToString(), truncated: true);
                }
                char character = value[index];
                if (!char.IsWhiteSpace(character)) {
                    if (builder.Length == MaximumScopeComparisonTextCharacters) {
                        return new ScopeComparisonText(builder.ToString(), truncated: true);
                    }
                    builder.Append(char.ToUpperInvariant(character));
                }
            }
        }

        return new ScopeComparisonText(builder.ToString(), truncated: false);
    }

    private enum ScopeRepresentation {
        NotRepresented,
        Represented,
        Incomplete
    }

    private readonly struct ScopeComparisonText {
        internal ScopeComparisonText(string value, bool truncated) {
            Value = value;
            Truncated = truncated;
        }

        internal string Value { get; }
        internal bool Truncated { get; }
    }

    private static bool ContainsOrdinal(string value, string candidate) {
#if NETSTANDARD2_0 || NETFRAMEWORK
        return value.IndexOf(candidate, StringComparison.Ordinal) >= 0;
#else
        return value.Contains(candidate, StringComparison.Ordinal);
#endif
    }

    /// <summary>
    /// Gets the maximum visible cell count across all logical table rows.
    /// </summary>
    /// <param name="table">Logical table to inspect.</param>
    /// <returns>The maximum row width, or zero when the table has no visible cells.</returns>
    public static int GetColumnCount(PdfLogicalTable table) =>
        GetColumnCount(table, CancellationToken.None);

    private static int GetColumnCount(PdfLogicalTable table, CancellationToken cancellationToken) {
        Guard.NotNull(table, nameof(table));

        int columnCount = 0;
        for (int i = 0; i < table.Rows.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            columnCount = Math.Max(columnCount, table.Rows[i].Count);
        }

        return columnCount;
    }

    /// <summary>
    /// Detects body columns whose non-empty cells look numeric and can be right-aligned by text emitters.
    /// </summary>
    /// <param name="table">Logical table to inspect.</param>
    /// <returns>A Boolean value per visible table column. True means all non-empty body cells in that column look numeric.</returns>
    public static bool[] DetectNumericColumns(PdfLogicalTable table) {
        Guard.NotNull(table, nameof(table));

        return DetectNumericColumns(table, GetColumnCount(table));
    }

    /// <summary>
    /// Detects extracted logical table columns whose non-empty cells can be safely converted to decimal values.
    /// </summary>
    /// <param name="data">Extracted logical table data to inspect.</param>
    /// <param name="culture">Preferred culture for localized numeric text. Invariant parsing is also attempted.</param>
    /// <returns>A Boolean value per extracted table column. True means every non-empty cell in that column parses as a decimal value.</returns>
    public static bool[] DetectParsableNumericColumns(PdfLogicalTableData data, CultureInfo? culture = null) {
        Guard.NotNull(data, nameof(data));

        var columns = new bool[data.Columns.Count];
        CultureInfo effectiveCulture = culture ?? CultureInfo.InvariantCulture;
        for (int columnIndex = 0; columnIndex < columns.Length; columnIndex++) {
            if (!data.IsNumericColumn(columnIndex)) {
                continue;
            }

            columns[columnIndex] = CanParseNumericColumn(data.Rows, columnIndex, effectiveCulture);
        }

        return columns;
    }

    /// <summary>
    /// Detects numeric body columns using a previously inferred table structure.
    /// </summary>
    /// <param name="table">Logical table to inspect.</param>
    /// <param name="structure">Inferred table structure that provides column count and body-row boundary.</param>
    /// <returns>A Boolean value per visible table column. True means all non-empty body cells in that column look numeric.</returns>
    public static bool[] DetectNumericColumns(PdfLogicalTable table, PdfLogicalTableStructure structure) {
        return DetectNumericColumns(table, structure, CancellationToken.None);
    }

    private static bool[] DetectNumericColumns(
        PdfLogicalTable table,
        PdfLogicalTableStructure structure,
        CancellationToken cancellationToken) {
        Guard.NotNull(table, nameof(table));
        Guard.NotNull(structure, nameof(structure));

        return DetectNumericColumns(
            table,
            structure.ColumnCount,
            structure.BodyStartRowIndex,
            cancellationToken);
    }

    /// <summary>
    /// Returns normalized logical body rows using a previously inferred table structure.
    /// </summary>
    /// <param name="table">Logical table to inspect.</param>
    /// <param name="structure">Inferred table structure that provides column count and body-row boundary.</param>
    /// <param name="maxRows">Maximum number of rows to return. Values less than or equal to zero return all body rows.</param>
    /// <returns>Body rows padded or trimmed to the inferred column count.</returns>
    public static IReadOnlyList<IReadOnlyList<string>> GetBodyRows(PdfLogicalTable table, PdfLogicalTableStructure structure, int maxRows = 0) {
        return GetBodyRows(table, structure, maxRows, CancellationToken.None);
    }

    private static IReadOnlyList<IReadOnlyList<string>> GetBodyRows(
        PdfLogicalTable table,
        PdfLogicalTableStructure structure,
        int maxRows,
        CancellationToken cancellationToken) {
        Guard.NotNull(table, nameof(table));
        Guard.NotNull(structure, nameof(structure));
        cancellationToken.ThrowIfCancellationRequested();

        int availableRows = Math.Max(0, table.Rows.Count - structure.BodyStartRowIndex);
        int rowCount = maxRows > 0 ? Math.Min(maxRows, availableRows) : availableRows;
        if (rowCount == 0 || structure.ColumnCount == 0) {
            return Array.Empty<IReadOnlyList<string>>();
        }

        var rows = new IReadOnlyList<string>[rowCount];
        for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            rows[rowIndex] = NormalizeRow(table.Rows[structure.BodyStartRowIndex + rowIndex], structure.ColumnCount);
        }

        return Array.AsReadOnly(rows);
    }

    internal static bool[] DetectNumericColumns(PdfLogicalTable table, int columnCount) {
        return DetectNumericColumns(table, columnCount, startRow: 1, CancellationToken.None);
    }

    private static bool[] DetectNumericColumns(
        PdfLogicalTable table,
        int columnCount,
        int startRow,
        CancellationToken cancellationToken) {
        var numericColumns = new bool[columnCount];
        if (table.Rows.Count <= startRow) {
            return numericColumns;
        }

        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            bool sawValue = false;
            bool allNumeric = true;
            for (int rowIndex = startRow; rowIndex < table.Rows.Count; rowIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                IReadOnlyList<string> row = table.Rows[rowIndex];
                string value = columnIndex < row.Count ? row[columnIndex] : string.Empty;
                if (string.IsNullOrWhiteSpace(value)) {
                    continue;
                }

                sawValue = true;
                if (!LooksLikeNumericValue(value)) {
                    allNumeric = false;
                    break;
                }
            }

            numericColumns[columnIndex] = sawValue && allNumeric;
        }

        return numericColumns;
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<string> NormalizeRow(IReadOnlyList<string> row, int columnCount) {
        var normalized = new string[columnCount];
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            normalized[columnIndex] = columnIndex < row.Count ? row[columnIndex] : string.Empty;
        }

        return Array.AsReadOnly(normalized);
    }

    /// <summary>
    /// Reports whether a table cell value looks numeric for Markdown and HTML alignment purposes.
    /// </summary>
    /// <param name="text">Cell text to inspect.</param>
    /// <returns>True when the value contains at least one digit and only numeric punctuation, whitespace, or currency symbols.</returns>
    public static bool LooksLikeNumericValue(string? text) {
        string value = text?.Trim() ?? string.Empty;
        if (value.Length == 0) {
            return false;
        }

        bool hasDigit = false;
        for (int index = 0; index < value.Length;) {
            int digit = CharUnicodeInfo.GetDecimalDigitValue(value, index);
            if (digit >= 0) {
                hasDigit = true;
                index += char.IsSurrogatePair(value, index) ? 2 : 1;
                continue;
            }

            UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(value, index);
            char c = value[index];
            if (char.IsWhiteSpace(c) || c == '.' || c == ',' || c == '-' || c == '+' || c == '(' || c == ')' || c == '%') {
                index++;
                continue;
            }

            if (category == UnicodeCategory.CurrencySymbol ||
                IsEquivalentNumericPunctuation(c)) {
                index += char.IsSurrogatePair(value, index) ? 2 : 1;
                continue;
            }

            return false;
        }

        return hasDigit;
    }

    /// <summary>
    /// Parses common invoice and statement numeric cell text into a decimal value for editable document exports.
    /// </summary>
    /// <param name="text">Cell text to parse.</param>
    /// <param name="culture">Preferred culture for localized numeric text. Invariant parsing is also attempted.</param>
    /// <param name="value">Parsed decimal value when parsing succeeds.</param>
    /// <returns>True when the text can be converted to a decimal value without treating percentages as ordinary numbers.</returns>
    public static bool TryParseNumericValue(string? text, CultureInfo? culture, out decimal value) {
        value = 0m;
        string source = text?.Trim() ?? string.Empty;
        if (source.Length == 0 || ContainsPercent(source)) {
            return false;
        }

        CultureInfo effectiveCulture = culture ?? CultureInfo.InvariantCulture;
        const NumberStyles styles = NumberStyles.Number | NumberStyles.AllowCurrencySymbol | NumberStyles.AllowParentheses;
        if (decimal.TryParse(source, styles, effectiveCulture, out value) ||
            decimal.TryParse(source, styles, CultureInfo.InvariantCulture, out value)) {
            return true;
        }

        if (!LooksLikeNumericValue(source)) {
            return false;
        }

        return TryParseNormalizedNumericValue(source, effectiveCulture, out value) ||
            (!ReferenceEquals(effectiveCulture, CultureInfo.InvariantCulture) &&
             TryParseNormalizedNumericValue(source, CultureInfo.InvariantCulture, out value));
    }

    private static bool CanParseNumericColumn(IReadOnlyList<IReadOnlyList<string>> rows, int columnIndex, CultureInfo culture) {
        bool sawValue = false;
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            IReadOnlyList<string> row = rows[rowIndex];
            string value = columnIndex < row.Count ? row[columnIndex] : string.Empty;
            if (string.IsNullOrWhiteSpace(value)) {
                continue;
            }

            if (!TryParseNumericValue(value, culture, out _)) {
                return false;
            }

            sawValue = true;
        }

        return sawValue;
    }

    private static bool ContainsPercent(string value) {
        for (int i = 0; i < value.Length; i++) {
            if (IsPercentSign(value[i])) {
                return true;
            }
        }

        return false;
    }

    private static bool TryParseNormalizedNumericValue(string source, CultureInfo culture, out decimal value) {
        value = 0m;
        bool negative = source.Length > 2 && IsOpenParenthesis(source[0]) && IsCloseParenthesis(source[source.Length - 1]);
        int start = negative ? 1 : 0;
        int end = negative ? source.Length - 1 : source.Length;
        var chars = new char[source.Length];
        int count = 0;
        for (int index = start; index < end;) {
            int digit = CharUnicodeInfo.GetDecimalDigitValue(source, index);
            if (digit >= 0) {
                chars[count++] = (char)('0' + digit);
                index += char.IsSurrogatePair(source, index) ? 2 : 1;
                continue;
            }

            char c = source[index];
            UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(source, index);
            int scalarLength = char.IsSurrogatePair(source, index) ? 2 : 1;
            index += scalarLength;
            if (char.IsWhiteSpace(c) || category == UnicodeCategory.CurrencySymbol) {
                continue;
            }
            switch (c) {
                case '-':
                case '+':
                    chars[count++] = c;
                    break;
                case '\u066B': // Arabic decimal separator
                    chars[count++] = '.';
                    break;
                case '\u066C': // Arabic thousands separator
                    break;
                case '.':
                case ',':
                case '\uFF0E': // fullwidth full stop
                case '\uFF0C': // fullwidth comma
                    char separator = c == '\uFF0E' ? '.' : c == '\uFF0C' ? ',' : c;
                    if (MatchesNumberSeparator(separator, culture.NumberFormat.NumberDecimalSeparator)) {
                        chars[count++] = '.';
                    } else if (!MatchesNumberSeparator(separator, culture.NumberFormat.NumberGroupSeparator)) {
                        return false;
                    }
                    break;
                case '\u2212':
                case '\uFE63':
                case '\uFF0D':
                    chars[count++] = '-';
                    break;
                case '\uFE62':
                case '\uFF0B':
                    chars[count++] = '+';
                    break;
                case '\'':
                case '\u02BC':
                case '\u2019':
                    break;
                default:
                    // Parentheses are valid only as one matched outer pair. Any other permitted
                    // punctuation is not part of the normalized numeric grammar.
                    return false;
            }
        }

        if (count == 0) {
            return false;
        }

        string normalized = new string(chars, 0, count);
        if (negative && normalized.Length > 0 && normalized[0] != '-') {
            normalized = "-" + normalized;
        }

        return decimal.TryParse(normalized, NumberStyles.Number, CultureInfo.InvariantCulture, out value);
    }

    private static bool MatchesNumberSeparator(char value, string separator) =>
        separator.Length == 1 && separator[0] == value;

    private static bool IsEquivalentNumericPunctuation(char value) =>
        value == '\u066B' || value == '\u066C' ||
        value == '\uFF0E' || value == '\uFF0C' ||
        value == '\u2212' || value == '\uFE63' || value == '\uFF0D' ||
        value == '\uFE62' || value == '\uFF0B' ||
        value == '\'' || value == '\u02BC' || value == '\u2019' ||
        IsOpenParenthesis(value) || IsCloseParenthesis(value) || IsPercentSign(value);

    private static bool IsOpenParenthesis(char value) => value == '(' || value == '\uFF08';

    private static bool IsCloseParenthesis(char value) => value == ')' || value == '\uFF09';

    internal static bool IsPercentSign(char value) =>
        value == '%' || value == '\u066A' || value == '\uFE6A' || value == '\uFF05';

    private static PdfInferenceEvidence[] BuildSchemaEvidence(PdfLogicalTable table, bool hasHeader) {
        if (!hasHeader) {
            return new[] { new PdfInferenceEvidence(
                "table.schema-unknown",
                "No tagged or structural evidence is strong enough to promote a row to table schema.",
                -0.5D) };
        }
        PdfInferenceEvidence? tagged = table.Evidence.FirstOrDefault(static evidence =>
            string.Equals(evidence.Code, "table.tagged-header-row", StringComparison.Ordinal));
        if (tagged is not null) return new[] { tagged };
        PdfInferenceEvidence? emphasis = table.Evidence.FirstOrDefault(static evidence =>
            string.Equals(evidence.Code, "table.header-emphasis", StringComparison.Ordinal));
        return emphasis is not null ? new[] { emphasis } : Array.Empty<PdfInferenceEvidence>();
    }

    private static string[] BuildUnnamedColumns(int columnCount) {
        return new string[columnCount];
    }
}
