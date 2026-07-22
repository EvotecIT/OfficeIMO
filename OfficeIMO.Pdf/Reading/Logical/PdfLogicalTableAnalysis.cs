using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Shared analysis helpers for logical PDF tables produced by <see cref="PdfLogicalDocument"/>.
/// </summary>
public static class PdfLogicalTableAnalysis {
    private const int DefaultMaximumScopeAnalysisComparisons = 10_000;
    private const int MaximumScopeComparisonTextCharacters = 512;
    private const int MaximumScopeSourceCharactersPerValue = 2048;
    /// <summary>
    /// Infers structured extraction metadata for a logical PDF table.
    /// </summary>
    /// <param name="table">Logical table to inspect.</param>
    /// <returns>Column names, body-row boundary, and table-shape flags for structured consumers.</returns>
    public static PdfLogicalTableStructure Analyze(PdfLogicalTable table) {
        Guard.NotNull(table, nameof(table));

        int columnCount = GetColumnCount(table);
        IReadOnlyList<string>? headerColumns = DetectHeaderColumns(table);
        bool hasHeader = headerColumns != null && headerColumns.Count == columnCount;
        bool isKeyValueTable = LooksLikeKeyValueTable(table);
        int bodyStartRowIndex = hasHeader ? 1 : 0;
        int totalBodyRowCount = Math.Max(0, table.Rows.Count - bodyStartRowIndex);
        IReadOnlyList<string> columns = hasHeader
            ? headerColumns!
            : isKeyValueTable
                ? KeyValueColumns
                : BuildFallbackColumns(columnCount);

        return new PdfLogicalTableStructure(
            columnCount,
            columns,
            bodyStartRowIndex,
            totalBodyRowCount,
            hasHeader,
            isKeyValueTable);
    }

    /// <summary>
    /// Extracts a normalized, structured table view for document readers and text emitters.
    /// </summary>
    /// <param name="table">Logical table to inspect.</param>
    /// <param name="maxRows">Maximum number of body rows to return. Values less than or equal to zero return all body rows.</param>
    /// <returns>Inferred columns, normalized body rows, numeric-column flags, and truncation metadata.</returns>
    public static PdfLogicalTableData Extract(PdfLogicalTable table, int maxRows = 0) {
        Guard.NotNull(table, nameof(table));

        PdfLogicalTableStructure structure = Analyze(table);
        IReadOnlyList<IReadOnlyList<string>> rows = GetBodyRows(table, structure, maxRows);
        IReadOnlyList<bool> numericColumns = DetectNumericColumns(table, structure);
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
    public static IReadOnlyList<PdfLogicalTableExtraction> ExtractTables(PdfLogicalDocument document, int maxRows = 0) {
        Guard.NotNull(document, nameof(document));

        return ExtractTables(document.Pages, maxRows);
    }

    /// <summary>
    /// Extracts normalized tables from the supplied logical pages in their current order.
    /// </summary>
    /// <param name="pages">Logical pages to inspect.</param>
    /// <param name="maxRows">Maximum number of body rows per table. Values less than or equal to zero return all body rows.</param>
    /// <returns>Page-aware normalized table extractions.</returns>
    public static IReadOnlyList<PdfLogicalTableExtraction> ExtractTables(IReadOnlyList<PdfLogicalPage> pages, int maxRows = 0) {
        Guard.NotNull(pages, nameof(pages));

        var extractions = new List<PdfLogicalTableExtraction>();
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            PdfLogicalPage page = pages[pageIndex];
            for (int tableIndex = 0; tableIndex < page.Tables.Count; tableIndex++) {
                PdfLogicalTable table = page.Tables[tableIndex];
                extractions.Add(new PdfLogicalTableExtraction(
                    pageIndex,
                    page.PageNumber,
                    tableIndex,
                    table,
                    Extract(table, maxRows)));
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

        return ExtractTables(new[] { page }, maxRows);
    }

    /// <summary>
    /// Describes the source-page content considered by table-only adapters.
    /// </summary>
    /// <param name="document">Logical PDF document to inspect.</param>
    /// <returns>
    /// Table counts plus visible and interactive page content that a table-only adapter will not import.
    /// </returns>
    public static PdfTableExtractionScopeReport AnalyzeExtractionScope(PdfLogicalDocument document) {
        return AnalyzeExtractionScope(document, DefaultMaximumScopeAnalysisComparisons);
    }

    /// <summary>
    /// Describes table extraction scope while bounding attacker-controlled text/table comparisons.
    /// </summary>
    public static PdfTableExtractionScopeReport AnalyzeExtractionScope(
        PdfLogicalDocument document,
        int maximumComparisons) {
        Guard.NotNull(document, nameof(document));
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
        int remainingComparisons = Math.Min(maximumComparisons, DefaultMaximumScopeAnalysisComparisons);
        bool analysisTruncated = false;
        var normalizedRows = new Dictionary<PdfLogicalTable, Dictionary<int, ScopeComparisonText>>();

        for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++) {
            PdfLogicalPage page = document.Pages[pageIndex];
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
        }

        return new PdfTableExtractionScopeReport(
            document.Pages.Count,
            pagesWithTables,
            detectedTableCount,
            nonTableTextBlockCount,
            vectorPrimitiveCount,
            imageCount,
            linkCount,
            formWidgetCount,
            annotationCount,
            pageActionCount,
            analysisTruncated);
    }

    /// <summary>
    /// Detects a simple header row from the first logical table row.
    /// </summary>
    /// <param name="table">Logical table to inspect.</param>
    /// <returns>Header column names when the first row looks like distinct text headers; otherwise null.</returns>
    public static IReadOnlyList<string>? DetectHeaderColumns(PdfLogicalTable table) {
        Guard.NotNull(table, nameof(table));

        int columnCount = GetColumnCount(table);
        if (table.Rows.Count <= 1 || columnCount <= 1) {
            return null;
        }

        IReadOnlyList<string> firstRow = table.Rows[0];
        if (firstRow.Count < columnCount) {
            return null;
        }

        var headers = new string[columnCount];
        var seenHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int nonNumericCount = 0;
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            string header = firstRow[columnIndex].Trim();
            if (header.Length == 0 || !seenHeaders.Add(header)) {
                return null;
            }

            if (!LooksLikeNumericValue(header)) {
                nonNumericCount++;
            }

            headers[columnIndex] = header;
        }

        if (columnCount == 2 &&
            !LooksLikeKeyValueHeader(headers) &&
            !IsValueHeader(headers[1]) &&
            LooksLikeHeaderlessKeyValueFirstRow(headers) &&
            LooksLikeKeyValueBody(table, startRow: 0)) {
            return null;
        }

        return nonNumericCount == columnCount ? headers : null;
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
    /// Reports whether a two-column logical table looks like label/value facts instead of a general data grid.
    /// </summary>
    /// <param name="table">Logical table to inspect.</param>
    /// <returns>True when the table has two visible columns and body rows look like non-numeric labels with values.</returns>
    public static bool LooksLikeKeyValueTable(PdfLogicalTable table) {
        Guard.NotNull(table, nameof(table));

        if (GetColumnCount(table) != 2 || table.Rows.Count == 0) {
            return false;
        }

        IReadOnlyList<string>? headerColumns = DetectHeaderColumns(table);
        bool hasHeader = headerColumns != null && headerColumns.Count == 2;
        if (hasHeader && !LooksLikeKeyValueHeader(headerColumns!)) {
            return false;
        }

        return LooksLikeKeyValueBody(table, hasHeader ? 1 : 0);
    }

    private static bool LooksLikeKeyValueBody(PdfLogicalTable table, int startRow) {
        int bodyRowCount = 0;
        for (int rowIndex = startRow; rowIndex < table.Rows.Count; rowIndex++) {
            IReadOnlyList<string> row = table.Rows[rowIndex];
            string key = row.Count > 0 ? row[0].Trim() : string.Empty;
            string value = row.Count > 1 ? row[1].Trim() : string.Empty;
            if (key.Length == 0 || value.Length == 0 || LooksLikeNumericValue(key)) {
                return false;
            }

            bodyRowCount++;
        }

        return bodyRowCount > 0;
    }

    /// <summary>
    /// Gets the maximum visible cell count across all logical table rows.
    /// </summary>
    /// <param name="table">Logical table to inspect.</param>
    /// <returns>The maximum row width, or zero when the table has no visible cells.</returns>
    public static int GetColumnCount(PdfLogicalTable table) {
        Guard.NotNull(table, nameof(table));

        int columnCount = 0;
        for (int i = 0; i < table.Rows.Count; i++) {
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
        Guard.NotNull(table, nameof(table));
        Guard.NotNull(structure, nameof(structure));

        return DetectNumericColumns(table, structure.ColumnCount, structure.BodyStartRowIndex);
    }

    /// <summary>
    /// Returns normalized logical body rows using a previously inferred table structure.
    /// </summary>
    /// <param name="table">Logical table to inspect.</param>
    /// <param name="structure">Inferred table structure that provides column count and body-row boundary.</param>
    /// <param name="maxRows">Maximum number of rows to return. Values less than or equal to zero return all body rows.</param>
    /// <returns>Body rows padded or trimmed to the inferred column count.</returns>
    public static IReadOnlyList<IReadOnlyList<string>> GetBodyRows(PdfLogicalTable table, PdfLogicalTableStructure structure, int maxRows = 0) {
        Guard.NotNull(table, nameof(table));
        Guard.NotNull(structure, nameof(structure));

        int availableRows = Math.Max(0, table.Rows.Count - structure.BodyStartRowIndex);
        int rowCount = maxRows > 0 ? Math.Min(maxRows, availableRows) : availableRows;
        if (rowCount == 0 || structure.ColumnCount == 0) {
            return Array.Empty<IReadOnlyList<string>>();
        }

        var rows = new IReadOnlyList<string>[rowCount];
        for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
            rows[rowIndex] = NormalizeRow(table.Rows[structure.BodyStartRowIndex + rowIndex], structure.ColumnCount);
        }

        return Array.AsReadOnly(rows);
    }

    internal static bool[] DetectNumericColumns(PdfLogicalTable table, int columnCount) {
        return DetectNumericColumns(table, columnCount, startRow: 1);
    }

    private static bool[] DetectNumericColumns(PdfLogicalTable table, int columnCount, int startRow) {
        var numericColumns = new bool[columnCount];
        if (table.Rows.Count <= startRow) {
            return numericColumns;
        }

        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            bool sawValue = false;
            bool allNumeric = true;
            for (int rowIndex = startRow; rowIndex < table.Rows.Count; rowIndex++) {
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
        for (int i = 0; i < value.Length; i++) {
            char c = value[i];
            if (char.IsDigit(c)) {
                hasDigit = true;
                continue;
            }

            if (char.IsWhiteSpace(c) || c == '.' || c == ',' || c == '-' || c == '+' || c == '(' || c == ')' || c == '%') {
                continue;
            }

            if (char.GetUnicodeCategory(c) == UnicodeCategory.CurrencySymbol) {
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

        return TryParseNormalizedNumericValue(source, out value);
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
            if (value[i] == '%') {
                return true;
            }
        }

        return false;
    }

    private static bool TryParseNormalizedNumericValue(string source, out decimal value) {
        value = 0m;
        bool negative = source.Length > 2 && source[0] == '(' && source[source.Length - 1] == ')';
        int start = negative ? 1 : 0;
        int end = negative ? source.Length - 1 : source.Length;
        var chars = new char[source.Length];
        int count = 0;
        for (int i = start; i < end; i++) {
            char c = source[i];
            if (char.IsDigit(c) || c == '.' || c == ',' || c == '-' || c == '+') {
                chars[count++] = c;
            }
        }

        if (count == 0) {
            return false;
        }

        string normalized = NormalizeNumberSeparators(new string(chars, 0, count));
        if (negative && normalized.Length > 0 && normalized[0] != '-') {
            normalized = "-" + normalized;
        }

        return decimal.TryParse(normalized, NumberStyles.Number, CultureInfo.InvariantCulture, out value);
    }

    private static string NormalizeNumberSeparators(string value) {
        int lastDot = value.LastIndexOf('.');
        int lastComma = value.LastIndexOf(',');
        char decimalSeparator = lastDot >= 0 && lastComma >= 0
            ? lastDot > lastComma ? '.' : ','
            : lastDot >= 0
                ? GetSingleSeparatorRole(value, '.') == NumberSeparatorRole.Decimal ? '.' : '\0'
                : lastComma >= 0 && GetSingleSeparatorRole(value, ',') == NumberSeparatorRole.Decimal ? ',' : '\0';

        var chars = new char[value.Length];
        int count = 0;
        for (int i = 0; i < value.Length; i++) {
            char c = value[i];
            if (c == '.' || c == ',') {
                if (c == decimalSeparator) {
                    chars[count++] = '.';
                }

                continue;
            }

            chars[count++] = c;
        }

        return new string(chars, 0, count);
    }

    private static NumberSeparatorRole GetSingleSeparatorRole(string value, char separator) {
        int first = value.IndexOf(separator);
        if (first < 0 || first != value.LastIndexOf(separator)) {
            return NumberSeparatorRole.Group;
        }

        int digitsAfter = 0;
        for (int i = first + 1; i < value.Length; i++) {
            if (char.IsDigit(value[i])) {
                digitsAfter++;
            }
        }

        return digitsAfter == 3 ? NumberSeparatorRole.Group : NumberSeparatorRole.Decimal;
    }

    private static bool LooksLikeKeyValueHeader(IReadOnlyList<string> headerColumns) {
        string keyHeader = headerColumns[0].Trim();
        string valueHeader = headerColumns[1].Trim();

        return IsKeyHeader(keyHeader) && IsValueHeader(valueHeader);
    }

    private static bool LooksLikeHeaderlessKeyValueFirstRow(string[] firstRow) {
        string key = firstRow[0].Trim();
        string value = firstRow[1].Trim();

        return !IsKeyHeader(key) && value.Any(static c => char.IsDigit(c));
    }

    private static bool IsKeyHeader(string header) {
        return string.Equals(header, "Key", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(header, "Field", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(header, "Item", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(header, "Label", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(header, "Name", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(header, "Property", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsValueHeader(string header) {
        return string.Equals(header, "Value", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(header, "Amount", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(header, "Price", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(header, "Qty", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(header, "Quantity", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(header, "Total", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(header, "Text", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(header, "Description", StringComparison.OrdinalIgnoreCase);
    }

    private static readonly IReadOnlyList<string> KeyValueColumns = new[] { "Key", "Value" };

    private enum NumberSeparatorRole {
        Group,
        Decimal
    }

    private static string[] BuildFallbackColumns(int columnCount) {
        return Enumerable.Range(1, columnCount)
            .Select(static column => "Column " + column.ToString(CultureInfo.InvariantCulture))
            .ToArray();
    }
}
