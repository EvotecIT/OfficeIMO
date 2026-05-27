using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

/// <summary>
/// Minimal, zero-dependency text extractor for simple PDFs produced by OfficeIMO.Pdf
/// and common external PDFs with basic text operators and common content-stream filters.
/// Not a general-purpose PDF parser; designed as a pragmatic starting point.
/// </summary>
public static class PdfTextExtractor {
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(2);
    private static readonly char[] SpaceSplitChars = new[] { ' ' };
    private static readonly char[] CsvQuoteChars = new[] { ',', '"', '\r', '\n' };
#if NET8_0_OR_GREATER
    private static readonly Regex ObjRegex = new Regex(@"(\d+)\s+0\s+obj", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex InfoRefRegex = new Regex(@"/Info\s+(\d+)\s+0\s+R", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex PageObjRegex = new Regex(@"<<(?:.*?)/Type\s*/Page\b(?:.*?)/Contents\s+(?:(?<single>\d+)\s+0\s+R|\[(?<array>[^\]]*)\])(?:.*?)/?>>", RegexOptions.Compiled | RegexOptions.Singleline | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex RefRegex = new Regex(@"(\d+)\s+0\s+R", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex StreamRegex = new Regex(@"stream\r?\n([\s\S]*?)\r?\nendstream", RegexOptions.Compiled | RegexOptions.Singleline | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex TjRegex = new Regex(@"\((?<txt>(?:\\.|[^\\\)])*)\)\s*Tj", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex HexTjRegex = new Regex(@"<(?<txt>[0-9A-Fa-f\s]+)>\s*Tj", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex QuoteLiteralRegex = new Regex(@"\((?<txt>(?:\\.|[^\\\)])*)\)\s*'", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex QuoteHexRegex = new Regex(@"<(?<txt>[0-9A-Fa-f\s]+)>\s*'", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex DoubleQuoteLiteralRegex = new Regex(@"(?<ws>[+-]?\d*\.?\d+)\s+(?<cs>[+-]?\d*\.?\d+)\s+\((?<txt>(?:\\.|[^\\\)])*)\)\s*""", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex DoubleQuoteHexRegex = new Regex(@"(?<ws>[+-]?\d*\.?\d+)\s+(?<cs>[+-]?\d*\.?\d+)\s+<(?<txt>[0-9A-Fa-f\s]+)>\s*""", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
#else
    private static readonly Regex ObjRegex = new Regex(@"(\d+)\s+0\s+obj", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex InfoRefRegex = new Regex(@"/Info\s+(\d+)\s+0\s+R", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex PageObjRegex = new Regex(@"<<(?:.|\n|\r)*?/Type\s*/Page\b(?:.|\n|\r)*?/Contents\s+(?:(?<single>\d+)\s+0\s+R|\[(?<array>[^\]]*)\])(?:.|\n|\r)*?>>", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex RefRegex = new Regex(@"(\d+)\s+0\s+R", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex StreamRegex = new Regex(@"stream\r?\n([\s\S]*?)\r?\nendstream", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex TjRegex = new Regex(@"\((?<txt>(?:\\.|[^\\\)])*)\)\s*Tj", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex HexTjRegex = new Regex(@"<(?<txt>[0-9A-Fa-f\s]+)>\s*Tj", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex QuoteLiteralRegex = new Regex(@"\((?<txt>(?:\\.|[^\\\)])*)\)\s*'", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex QuoteHexRegex = new Regex(@"<(?<txt>[0-9A-Fa-f\s]+)>\s*'", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex DoubleQuoteLiteralRegex = new Regex(@"(?<ws>[+-]?\d*\.?\d+)\s+(?<cs>[+-]?\d*\.?\d+)\s+\((?<txt>(?:\\.|[^\\\)])*)\)\s*""", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex DoubleQuoteHexRegex = new Regex(@"(?<ws>[+-]?\d*\.?\d+)\s+(?<cs>[+-]?\d*\.?\d+)\s+<(?<txt>[0-9A-Fa-f\s]+)>\s*""", RegexOptions.Compiled, RegexTimeout);
#endif

    /// <summary>Extracts plain text from all pages, concatenated with blank lines between pages.</summary>
    public static string ExtractAllText(string path) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        var bytes = File.ReadAllBytes(path);
        return ExtractAllText(bytes);
    }

    /// <summary>Extracts plain text from all pages using layout options such as column detection and header/footer trimming.</summary>
    public static string ExtractAllText(string path, PdfTextLayoutOptions? options) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        if (options is null) {
            return ExtractAllText(path);
        }

        return PdfReadDocument.Load(path).ExtractTextWithColumns(options);
    }

    /// <summary>Extracts plain text from all pages and writes UTF-8 text to <paramref name="outputStream"/>.</summary>
    public static void ExtractAllText(string inputPath, Stream outputStream) {
        ExtractAllText(inputPath, outputStream, null);
    }

    /// <summary>Extracts plain text from all pages using layout options and writes UTF-8 text to <paramref name="outputStream"/>.</summary>
    public static void ExtractAllText(string inputPath, Stream outputStream, PdfTextLayoutOptions? options) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteTextOutput(outputStream, ExtractAllText(inputPath, options));
    }

    /// <summary>Extracts plain text from all pages and writes UTF-8 text to <paramref name="outputPath"/>.</summary>
    public static void ExtractAllText(string inputPath, string outputPath) {
        ExtractAllText(inputPath, outputPath, null);
    }

    /// <summary>Extracts plain text from all pages using layout options and writes UTF-8 text to <paramref name="outputPath"/>.</summary>
    public static void ExtractAllText(string inputPath, string outputPath, PdfTextLayoutOptions? options) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        string fullOutputPath = ValidateOutputPath(outputPath);

        WriteTextOutput(fullOutputPath, ExtractAllText(inputPath, options));
    }

    /// <summary>Extracts plain text from each page in document order.</summary>
    public static IReadOnlyList<string> ExtractTextByPage(string path) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractTextByPage(PdfReadDocument.Load(path));
    }

    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<string> ExtractTextByPageRanges(string path, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractTextByPageRanges(PdfReadDocument.Load(path), pageRanges);
    }

    /// <summary>Extracts structured content for each page, including detected lines, lists, leader rows, and simple tables.</summary>
    public static IReadOnlyList<StructuredPage> ExtractStructuredByPage(string path, PdfTextLayoutOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return PdfReadDocument.Load(path).ExtractStructuredPages(options);
    }

    /// <summary>Extracts structured content from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredPage> ExtractStructuredByPageRanges(string path, params PdfPageRange[] pageRanges) {
        return ExtractStructuredByPageRanges(path, (PdfTextLayoutOptions?)null, pageRanges);
    }

    /// <summary>Extracts structured content from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredPage> ExtractStructuredByPageRanges(string path, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractStructuredByPageRanges(PdfReadDocument.Load(path), options, pageRanges);
    }

    /// <summary>Extracts detected paragraphs grouped by page while preserving paragraph geometry.</summary>
    public static IReadOnlyList<StructuredParagraphPage> ExtractParagraphsByPage(string path, PdfTextLayoutOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return PdfReadDocument.Load(path).ExtractParagraphsByPage(options);
    }

    /// <summary>Extracts detected paragraphs from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredParagraphPage> ExtractParagraphsByPageRanges(string path, params PdfPageRange[] pageRanges) {
        return ExtractParagraphsByPageRanges(path, (PdfTextLayoutOptions?)null, pageRanges);
    }

    /// <summary>Extracts detected paragraphs from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredParagraphPage> ExtractParagraphsByPageRanges(string path, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractParagraphsByPageRanges(PdfReadDocument.Load(path), options, pageRanges);
    }

    /// <summary>Extracts detected headings grouped by page while preserving heading geometry.</summary>
    public static IReadOnlyList<StructuredHeadingPage> ExtractHeadingsByPage(string path, PdfTextLayoutOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return PdfReadDocument.Load(path).ExtractHeadingsByPage(options);
    }

    /// <summary>Extracts detected headings from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredHeadingPage> ExtractHeadingsByPageRanges(string path, params PdfPageRange[] pageRanges) {
        return ExtractHeadingsByPageRanges(path, (PdfTextLayoutOptions?)null, pageRanges);
    }

    /// <summary>Extracts detected headings from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredHeadingPage> ExtractHeadingsByPageRanges(string path, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractHeadingsByPageRanges(PdfReadDocument.Load(path), options, pageRanges);
    }

    /// <summary>Extracts detected list items grouped by page while preserving marker and nesting hints.</summary>
    public static IReadOnlyList<StructuredListItemPage> ExtractListItemsByPage(string path, PdfTextLayoutOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return PdfReadDocument.Load(path).ExtractListItemsByPage(options);
    }

    /// <summary>Extracts detected list items from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredListItemPage> ExtractListItemsByPageRanges(string path, params PdfPageRange[] pageRanges) {
        return ExtractListItemsByPageRanges(path, (PdfTextLayoutOptions?)null, pageRanges);
    }

    /// <summary>Extracts detected list items from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredListItemPage> ExtractListItemsByPageRanges(string path, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractListItemsByPageRanges(PdfReadDocument.Load(path), options, pageRanges);
    }

    /// <summary>Extracts detected tables grouped by page while preserving table geometry.</summary>
    public static IReadOnlyList<StructuredTablePage> ExtractTablesByPage(string path, PdfTextLayoutOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return PdfReadDocument.Load(path).ExtractTablesByPage(options);
    }

    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredTablePage> ExtractTablesByPageRanges(string path, params PdfPageRange[] pageRanges) {
        return ExtractTablesByPageRanges(path, (PdfTextLayoutOptions?)null, pageRanges);
    }

    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredTablePage> ExtractTablesByPageRanges(string path, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractTablesByPageRanges(PdfReadDocument.Load(path), options, pageRanges);
    }

    /// <summary>Extracts detected tables and writes one CSV file per detected table.</summary>
    public static IReadOnlyList<string> ExtractTablesByPage(string inputPath, string outputDirectory, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(inputPath, nameof(inputPath));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var tablePages = ExtractTablesByPage(inputPath, options);
        return WriteTableCsvFiles(inputPath, fullOutputDirectory, tablePages);
    }

    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges and writes one CSV file per detected table.</summary>
    public static IReadOnlyList<string> ExtractTablesByPageRanges(string inputPath, string outputDirectory, params PdfPageRange[] pageRanges) {
        return ExtractTablesByPageRanges(inputPath, outputDirectory, (PdfTextLayoutOptions?)null, pageRanges);
    }

    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges and writes one CSV file per detected table.</summary>
    public static IReadOnlyList<string> ExtractTablesByPageRanges(string inputPath, string outputDirectory, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(inputPath, nameof(inputPath));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var tablePages = ExtractTablesByPageRanges(inputPath, options, pageRanges);
        return WriteTableCsvFiles(inputPath, fullOutputDirectory, tablePages);
    }

    /// <summary>Extracts detected tables from the current stream position and writes one CSV file per detected table.</summary>
    public static IReadOnlyList<string> ExtractTablesByPage(Stream stream, string outputDirectory, string baseName = "table", PdfTextLayoutOptions? options = null) {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var tablePages = ExtractTablesByPage(stream, options);
        return WriteTableCsvFiles(baseName, fullOutputDirectory, tablePages);
    }

    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges from the current stream position and writes one CSV file per detected table.</summary>
    public static IReadOnlyList<string> ExtractTablesByPageRanges(Stream stream, string outputDirectory, string baseName = "table", params PdfPageRange[] pageRanges) {
        return ExtractTablesByPageRanges(stream, outputDirectory, baseName, (PdfTextLayoutOptions?)null, pageRanges);
    }

    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges from the current stream position and writes one CSV file per detected table.</summary>
    public static IReadOnlyList<string> ExtractTablesByPageRanges(Stream stream, string outputDirectory, string baseName, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var tablePages = ExtractTablesByPageRanges(stream, options, pageRanges);
        return WriteTableCsvFiles(baseName, fullOutputDirectory, tablePages);
    }

    /// <summary>Extracts detected tables from bytes and writes one CSV file per detected table.</summary>
    public static IReadOnlyList<string> ExtractTablesByPage(byte[] pdf, string outputDirectory, string baseName = "table", PdfTextLayoutOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var tablePages = ExtractTablesByPage(pdf, options);
        return WriteTableCsvFiles(baseName, fullOutputDirectory, tablePages);
    }

    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges from bytes and writes one CSV file per detected table.</summary>
    public static IReadOnlyList<string> ExtractTablesByPageRanges(byte[] pdf, string outputDirectory, string baseName = "table", params PdfPageRange[] pageRanges) {
        return ExtractTablesByPageRanges(pdf, outputDirectory, baseName, (PdfTextLayoutOptions?)null, pageRanges);
    }

    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges from bytes and writes one CSV file per detected table.</summary>
    public static IReadOnlyList<string> ExtractTablesByPageRanges(byte[] pdf, string outputDirectory, string baseName, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var tablePages = ExtractTablesByPageRanges(pdf, options, pageRanges);
        return WriteTableCsvFiles(baseName, fullOutputDirectory, tablePages);
    }

    /// <summary>Extracts plain text from each page using layout options such as column detection and header/footer trimming.</summary>
    public static IReadOnlyList<string> ExtractTextByPage(string path, PdfTextLayoutOptions? options) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        if (options is null) {
            return ExtractTextByPage(path);
        }

        return ExtractTextByPage(PdfReadDocument.Load(path), options);
    }

    /// <summary>Extracts plain text from each page and writes one UTF-8 text file per page.</summary>
    public static IReadOnlyList<string> ExtractTextByPage(string inputPath, string outputDirectory) {
        Guard.NotNull(inputPath, nameof(inputPath));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractTextByPage(inputPath);
        return WriteTextPages(inputPath, fullOutputDirectory, pages);
    }

    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges and writes one UTF-8 text file per selected source page.</summary>
    public static IReadOnlyList<string> ExtractTextByPageRanges(string inputPath, string outputDirectory, params PdfPageRange[] pageRanges) {
        Guard.NotNull(inputPath, nameof(inputPath));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractSelectedTextPages(PdfReadDocument.Load(inputPath), pageRanges);
        return WriteTextPages(inputPath, fullOutputDirectory, pages);
    }

    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges with layout options and writes one UTF-8 text file per selected source page.</summary>
    public static IReadOnlyList<string> ExtractTextByPageRanges(string inputPath, string outputDirectory, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(inputPath, nameof(inputPath));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractSelectedTextPages(PdfReadDocument.Load(inputPath), options, pageRanges);
        return WriteTextPages(inputPath, fullOutputDirectory, pages);
    }

    /// <summary>Extracts plain text from each page with layout options and writes one UTF-8 text file per page.</summary>
    public static IReadOnlyList<string> ExtractTextByPage(string inputPath, string outputDirectory, PdfTextLayoutOptions? options) {
        Guard.NotNull(inputPath, nameof(inputPath));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractTextByPage(inputPath, options);
        return WriteTextPages(inputPath, fullOutputDirectory, pages);
    }

    private static List<string> WriteTextPages(string baseName, string fullOutputDirectory, IReadOnlyList<string> pages) {
        string safeBaseName = GetSafeBaseName(baseName, "page");

        var paths = new List<string>(pages.Count);
        for (int i = 0; i < pages.Count; i++) {
            string outputPath = Path.Combine(fullOutputDirectory, safeBaseName + "-page-" + (i + 1).ToString("0000", System.Globalization.CultureInfo.InvariantCulture) + ".txt");
            File.WriteAllText(outputPath, pages[i], Encoding.UTF8);
            paths.Add(outputPath);
        }

        return paths;
    }

    private static List<string> WriteTextPages(string baseName, string fullOutputDirectory, IReadOnlyList<SelectedTextPage> pages) {
        string safeBaseName = GetSafeBaseName(baseName, "page");

        var paths = new List<string>(pages.Count);
        var pageOccurrences = new Dictionary<int, int>();
        for (int i = 0; i < pages.Count; i++) {
            int occurrence = IncrementOccurrence(pageOccurrences, pages[i].PageNumber);
            string outputPath = Path.Combine(
                fullOutputDirectory,
                safeBaseName +
                "-page-" + pages[i].PageNumber.ToString("0000", System.Globalization.CultureInfo.InvariantCulture) +
                BuildOccurrenceSuffix(occurrence) +
                ".txt");
            File.WriteAllText(outputPath, pages[i].Text, Encoding.UTF8);
            paths.Add(outputPath);
        }

        return paths;
    }

    private static List<string> WriteTableCsvFiles(string baseName, string fullOutputDirectory, IReadOnlyList<StructuredTablePage> tablePages) {
        string safeBaseName = GetSafeBaseName(baseName, "table");

        var paths = new List<string>();
        var pageOccurrences = new Dictionary<int, int>();
        foreach (var page in tablePages) {
            int occurrence = IncrementOccurrence(pageOccurrences, page.PageNumber);
            for (int tableIndex = 0; tableIndex < page.Tables.Count; tableIndex++) {
                string outputPath = Path.Combine(
                    fullOutputDirectory,
                    safeBaseName +
                    "-page-" + page.PageNumber.ToString("0000", System.Globalization.CultureInfo.InvariantCulture) +
                    BuildOccurrenceSuffix(occurrence) +
                    "-table-" + (tableIndex + 1).ToString("0000", System.Globalization.CultureInfo.InvariantCulture) +
                    ".csv");

                File.WriteAllText(outputPath, BuildCsv(page.Tables[tableIndex]), Encoding.UTF8);
                paths.Add(outputPath);
            }
        }

        return paths;
    }

    private static int IncrementOccurrence(Dictionary<int, int> occurrences, int key) {
        occurrences.TryGetValue(key, out int occurrence);
        occurrence++;
        occurrences[key] = occurrence;
        return occurrence;
    }

    private static string BuildOccurrenceSuffix(int occurrence) {
        return occurrence <= 1
            ? string.Empty
            : "-occurrence-" + occurrence.ToString("0000", System.Globalization.CultureInfo.InvariantCulture);
    }

    private static string GetSafeBaseName(string? baseName, string fallback) {
        string safeBaseName = Path.GetFileNameWithoutExtension(baseName ?? string.Empty) ?? string.Empty;
        return string.IsNullOrWhiteSpace(safeBaseName) ? fallback : safeBaseName;
    }

    private static string BuildCsv(StructuredTable table) {
        var sb = new StringBuilder();
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            var row = table.Rows[rowIndex];
            for (int cellIndex = 0; cellIndex < row.Length; cellIndex++) {
                if (cellIndex > 0) {
                    sb.Append(',');
                }

                sb.Append(EscapeCsvCell(row[cellIndex]));
            }

            if (rowIndex + 1 < table.Rows.Count) {
                sb.AppendLine();
            }
        }

        return sb.ToString();
    }

    private static string EscapeCsvCell(string? value) {
        string cell = value ?? string.Empty;
        if (cell.Length == 0) {
            return string.Empty;
        }

        bool quote = cell.IndexOfAny(CsvQuoteChars) >= 0;
        if (!quote) {
            return cell;
        }

        return "\"" + cell.Replace("\"", "\"\"") + "\"";
    }

    private static string ValidateOutputDirectory(string outputDirectory) {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
        if (string.IsNullOrWhiteSpace(outputDirectory)) {
            throw new ArgumentException("Output directory cannot be empty or whitespace.", nameof(outputDirectory));
        }

        string fullOutputDirectory;
        try {
            fullOutputDirectory = Path.GetFullPath(outputDirectory);
        } catch (Exception ex) {
            throw new ArgumentException("Output directory is invalid.", nameof(outputDirectory), ex);
        }

        if (File.Exists(fullOutputDirectory)) {
            throw new ArgumentException("Output directory refers to a file; a directory path is required.", nameof(outputDirectory));
        }

        Directory.CreateDirectory(fullOutputDirectory);
        return fullOutputDirectory;
    }

    private static void WriteTextOutput(Stream outputStream, string text) {
        ValidateWritableOutputStream(outputStream);
        byte[] bytes = new UTF8Encoding(false).GetBytes(text);
        outputStream.Write(bytes, 0, bytes.Length);
    }

    private static void WriteTextOutput(string outputPath, string text) {
        string fullPath = ValidateOutputPath(outputPath);
        var directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        File.WriteAllText(fullPath, text, new UTF8Encoding(false));
    }

    private static void ValidateWritableOutputStream(Stream outputStream) {
        Guard.NotNull(outputStream, nameof(outputStream));
        if (!outputStream.CanWrite) {
            throw new ArgumentException("Stream must be writable.", nameof(outputStream));
        }
    }

    private static string ValidateOutputPath(string outputPath) {
        Guard.NotNull(outputPath, nameof(outputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) {
            throw new ArgumentException("Output path cannot be empty or whitespace.", nameof(outputPath));
        }

        string fullPath;
        try {
            fullPath = Path.GetFullPath(outputPath);
        } catch (Exception ex) {
            throw new ArgumentException("Output path is invalid.", nameof(outputPath), ex);
        }

        if (Directory.Exists(fullPath) && (File.GetAttributes(fullPath) & FileAttributes.Directory) == FileAttributes.Directory) {
            throw new ArgumentException("Output path refers to a directory; a file path is required.", nameof(outputPath));
        }

        var fileName = Path.GetFileName(fullPath);
        if (string.IsNullOrEmpty(fileName)) {
            throw new ArgumentException("Output path must include a file name.", nameof(outputPath));
        }

        if (fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) {
            throw new ArgumentException("Output path contains invalid file name characters.", nameof(outputPath));
        }

        return fullPath;
    }

    /// <summary>Extracts plain text from all pages from the current position of a readable stream.</summary>
    public static string ExtractAllText(Stream stream) {
        return ExtractAllText(ReadAllBytes(stream));
    }

    /// <summary>Extracts plain text from all pages from the current stream position using layout options such as column detection and header/footer trimming.</summary>
    public static string ExtractAllText(Stream stream, PdfTextLayoutOptions? options) {
        if (options is null) {
            return ExtractAllText(stream);
        }

        return PdfReadDocument.Load(stream).ExtractTextWithColumns(options);
    }

    /// <summary>Extracts plain text from all pages from the current position of a readable stream and writes UTF-8 text to <paramref name="outputStream"/>.</summary>
    public static void ExtractAllText(Stream inputStream, Stream outputStream) {
        ExtractAllText(inputStream, outputStream, null);
    }

    /// <summary>Extracts plain text from all pages from the current position of a readable stream using layout options and writes UTF-8 text to <paramref name="outputStream"/>.</summary>
    public static void ExtractAllText(Stream inputStream, Stream outputStream, PdfTextLayoutOptions? options) {
        ValidateWritableOutputStream(outputStream);
        WriteTextOutput(outputStream, ExtractAllText(inputStream, options));
    }

    /// <summary>Extracts plain text from all pages from the current position of a readable stream and writes UTF-8 text to <paramref name="outputPath"/>.</summary>
    public static void ExtractAllText(Stream inputStream, string outputPath) {
        ExtractAllText(inputStream, outputPath, null);
    }

    /// <summary>Extracts plain text from all pages from the current position of a readable stream using layout options and writes UTF-8 text to <paramref name="outputPath"/>.</summary>
    public static void ExtractAllText(Stream inputStream, string outputPath, PdfTextLayoutOptions? options) {
        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteTextOutput(fullOutputPath, ExtractAllText(inputStream, options));
    }

    /// <summary>Extracts plain text from each page from the current position of a readable stream.</summary>
    public static IReadOnlyList<string> ExtractTextByPage(Stream stream) {
        return ExtractTextByPage(PdfReadDocument.Load(stream));
    }

    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges from the current position of a readable stream.</summary>
    public static IReadOnlyList<string> ExtractTextByPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return ExtractTextByPageRanges(PdfReadDocument.Load(stream), pageRanges);
    }

    /// <summary>Extracts plain text from each page from the current stream position and writes one UTF-8 text file per page.</summary>
    public static IReadOnlyList<string> ExtractTextByPage(Stream stream, string outputDirectory, string baseName = "page", PdfTextLayoutOptions? options = null) {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractTextByPage(stream, options);
        return WriteTextPages(baseName, fullOutputDirectory, pages);
    }

    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges from the current stream position and writes one UTF-8 text file per selected source page.</summary>
    public static IReadOnlyList<string> ExtractTextByPageRanges(Stream stream, string outputDirectory, string baseName = "page", params PdfPageRange[] pageRanges) {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractSelectedTextPages(PdfReadDocument.Load(stream), pageRanges);
        return WriteTextPages(baseName, fullOutputDirectory, pages);
    }

    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges from the current stream position with layout options and writes one UTF-8 text file per selected source page.</summary>
    public static IReadOnlyList<string> ExtractTextByPageRanges(Stream stream, string outputDirectory, string baseName, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractSelectedTextPages(PdfReadDocument.Load(stream), options, pageRanges);
        return WriteTextPages(baseName, fullOutputDirectory, pages);
    }

    /// <summary>Extracts plain text from each page from bytes and writes one UTF-8 text file per page.</summary>
    public static IReadOnlyList<string> ExtractTextByPage(byte[] pdf, string outputDirectory, string baseName = "page", PdfTextLayoutOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractTextByPage(pdf, options);
        return WriteTextPages(baseName, fullOutputDirectory, pages);
    }

    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges from bytes and writes one UTF-8 text file per selected source page.</summary>
    public static IReadOnlyList<string> ExtractTextByPageRanges(byte[] pdf, string outputDirectory, string baseName = "page", params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractSelectedTextPages(PdfReadDocument.Load(pdf), pageRanges);
        return WriteTextPages(baseName, fullOutputDirectory, pages);
    }

    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges from bytes with layout options and writes one UTF-8 text file per selected source page.</summary>
    public static IReadOnlyList<string> ExtractTextByPageRanges(byte[] pdf, string outputDirectory, string baseName, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractSelectedTextPages(PdfReadDocument.Load(pdf), options, pageRanges);
        return WriteTextPages(baseName, fullOutputDirectory, pages);
    }

    /// <summary>Extracts structured content for each page from the current stream position.</summary>
    public static IReadOnlyList<StructuredPage> ExtractStructuredByPage(Stream stream, PdfTextLayoutOptions? options = null) {
        return PdfReadDocument.Load(stream).ExtractStructuredPages(options);
    }

    /// <summary>Extracts structured content from the supplied inclusive one-based page ranges from the current stream position.</summary>
    public static IReadOnlyList<StructuredPage> ExtractStructuredByPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return ExtractStructuredByPageRanges(stream, (PdfTextLayoutOptions?)null, pageRanges);
    }

    /// <summary>Extracts structured content from the supplied inclusive one-based page ranges from the current stream position.</summary>
    public static IReadOnlyList<StructuredPage> ExtractStructuredByPageRanges(Stream stream, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        return ExtractStructuredByPageRanges(PdfReadDocument.Load(stream), options, pageRanges);
    }

    /// <summary>Extracts detected paragraphs grouped by page from the current stream position.</summary>
    public static IReadOnlyList<StructuredParagraphPage> ExtractParagraphsByPage(Stream stream, PdfTextLayoutOptions? options = null) {
        return PdfReadDocument.Load(stream).ExtractParagraphsByPage(options);
    }

    /// <summary>Extracts detected paragraphs from the supplied inclusive one-based page ranges from the current stream position.</summary>
    public static IReadOnlyList<StructuredParagraphPage> ExtractParagraphsByPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return ExtractParagraphsByPageRanges(stream, (PdfTextLayoutOptions?)null, pageRanges);
    }

    /// <summary>Extracts detected paragraphs from the supplied inclusive one-based page ranges from the current stream position.</summary>
    public static IReadOnlyList<StructuredParagraphPage> ExtractParagraphsByPageRanges(Stream stream, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        return ExtractParagraphsByPageRanges(PdfReadDocument.Load(stream), options, pageRanges);
    }

    /// <summary>Extracts detected headings grouped by page from the current stream position.</summary>
    public static IReadOnlyList<StructuredHeadingPage> ExtractHeadingsByPage(Stream stream, PdfTextLayoutOptions? options = null) {
        return PdfReadDocument.Load(stream).ExtractHeadingsByPage(options);
    }

    /// <summary>Extracts detected headings from the supplied inclusive one-based page ranges from the current stream position.</summary>
    public static IReadOnlyList<StructuredHeadingPage> ExtractHeadingsByPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return ExtractHeadingsByPageRanges(stream, (PdfTextLayoutOptions?)null, pageRanges);
    }

    /// <summary>Extracts detected headings from the supplied inclusive one-based page ranges from the current stream position.</summary>
    public static IReadOnlyList<StructuredHeadingPage> ExtractHeadingsByPageRanges(Stream stream, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        return ExtractHeadingsByPageRanges(PdfReadDocument.Load(stream), options, pageRanges);
    }

    /// <summary>Extracts detected list items grouped by page from the current stream position.</summary>
    public static IReadOnlyList<StructuredListItemPage> ExtractListItemsByPage(Stream stream, PdfTextLayoutOptions? options = null) {
        return PdfReadDocument.Load(stream).ExtractListItemsByPage(options);
    }

    /// <summary>Extracts detected list items from the supplied inclusive one-based page ranges from the current stream position.</summary>
    public static IReadOnlyList<StructuredListItemPage> ExtractListItemsByPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return ExtractListItemsByPageRanges(stream, (PdfTextLayoutOptions?)null, pageRanges);
    }

    /// <summary>Extracts detected list items from the supplied inclusive one-based page ranges from the current stream position.</summary>
    public static IReadOnlyList<StructuredListItemPage> ExtractListItemsByPageRanges(Stream stream, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        return ExtractListItemsByPageRanges(PdfReadDocument.Load(stream), options, pageRanges);
    }

    /// <summary>Extracts detected tables grouped by page from the current stream position.</summary>
    public static IReadOnlyList<StructuredTablePage> ExtractTablesByPage(Stream stream, PdfTextLayoutOptions? options = null) {
        return PdfReadDocument.Load(stream).ExtractTablesByPage(options);
    }

    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges from the current stream position.</summary>
    public static IReadOnlyList<StructuredTablePage> ExtractTablesByPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return ExtractTablesByPageRanges(stream, (PdfTextLayoutOptions?)null, pageRanges);
    }

    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges from the current stream position.</summary>
    public static IReadOnlyList<StructuredTablePage> ExtractTablesByPageRanges(Stream stream, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        return ExtractTablesByPageRanges(PdfReadDocument.Load(stream), options, pageRanges);
    }

    /// <summary>Extracts plain text from each page from the current stream position using layout options such as column detection and header/footer trimming.</summary>
    public static IReadOnlyList<string> ExtractTextByPage(Stream stream, PdfTextLayoutOptions? options) {
        if (options is null) {
            return ExtractTextByPage(stream);
        }

        return ExtractTextByPage(PdfReadDocument.Load(stream), options);
    }

    /// <summary>Extracts plain text from all pages, concatenated with blank lines between pages.</summary>
    public static string ExtractAllText(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string? readModelText = null;
        try {
            readModelText = PdfReadDocument.Load(pdf).ExtractText();
        } catch (Exception ex) when (ex is not NotSupportedException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
            // Keep the legacy stream scan as a fallback for malformed-but-readable PDFs.
        }

        var (parsedObjects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        var map = BuildObjectMap(pdf, out _);
        var pages = CollectPages(parsedObjects, trailerRaw);
        var sb = new StringBuilder();

        if (pages.Count > 0) {
            for (int i = 0; i < pages.Count; i++) {
                string pageText = ExtractTextFromPage(pages[i], parsedObjects, map);
                if (string.IsNullOrWhiteSpace(pageText)) {
                    continue;
                }

                if (sb.Length > 0) {
                    sb.AppendLine();
                }
                sb.Append(pageText);
            }

            if (sb.Length > 0) {
                return ChooseAllText(readModelText, sb.ToString());
            }
        }

        var pageContents = FindPageContentIds(pdf);
        for (int i = 0; i < pageContents.Count; i++) {
            var pageText = new StringBuilder();
            foreach (int contentId in pageContents[i]) {
                if (TryGetContentStreamContent(parsedObjects, map, contentId, out string content)) {
                    pageText.Append(ExtractTextFromContentStream(content));
                }
            }

            if (pageText.Length == 0) {
                continue;
            }

            if (sb.Length > 0) {
                sb.AppendLine();
            }
            sb.Append(pageText);
        }
        return ChooseAllText(readModelText, sb.ToString());
    }

    private static string ChooseAllText(string? readModelText, string legacyText) {
        if (string.IsNullOrWhiteSpace(legacyText)) {
            return readModelText ?? string.Empty;
        }

        if (string.IsNullOrWhiteSpace(readModelText)) {
            return legacyText;
        }

        string readableText = readModelText!;
        return CountTextSeparators(readableText) > CountTextSeparators(legacyText)
            ? readableText
            : legacyText;
    }

    private static int CountTextSeparators(string value) {
        int count = 0;
        for (int i = 0; i < value.Length; i++) {
            if (char.IsWhiteSpace(value[i])) {
                count++;
            }
        }

        return count;
    }

    /// <summary>Extracts plain text from all pages using layout options such as column detection and header/footer trimming.</summary>
    public static string ExtractAllText(byte[] pdf, PdfTextLayoutOptions? options) {
        Guard.NotNull(pdf, nameof(pdf));
        if (options is null) {
            return ExtractAllText(pdf);
        }

        return PdfReadDocument.Load(pdf).ExtractTextWithColumns(options);
    }

    /// <summary>Extracts plain text from all pages and writes UTF-8 text to <paramref name="outputStream"/>.</summary>
    public static void ExtractAllText(byte[] pdf, Stream outputStream) {
        ExtractAllText(pdf, outputStream, null);
    }

    /// <summary>Extracts plain text from all pages using layout options and writes UTF-8 text to <paramref name="outputStream"/>.</summary>
    public static void ExtractAllText(byte[] pdf, Stream outputStream, PdfTextLayoutOptions? options) {
        Guard.NotNull(pdf, nameof(pdf));
        ValidateWritableOutputStream(outputStream);

        WriteTextOutput(outputStream, ExtractAllText(pdf, options));
    }

    /// <summary>Extracts plain text from all pages and writes UTF-8 text to <paramref name="outputPath"/>.</summary>
    public static void ExtractAllText(byte[] pdf, string outputPath) {
        ExtractAllText(pdf, outputPath, null);
    }

    /// <summary>Extracts plain text from all pages using layout options and writes UTF-8 text to <paramref name="outputPath"/>.</summary>
    public static void ExtractAllText(byte[] pdf, string outputPath, PdfTextLayoutOptions? options) {
        Guard.NotNull(pdf, nameof(pdf));
        string fullOutputPath = ValidateOutputPath(outputPath);

        WriteTextOutput(fullOutputPath, ExtractAllText(pdf, options));
    }

    /// <summary>Extracts plain text from each page in document order.</summary>
    public static IReadOnlyList<string> ExtractTextByPage(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractTextByPage(PdfReadDocument.Load(pdf));
    }

    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<string> ExtractTextByPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractTextByPageRanges(PdfReadDocument.Load(pdf), pageRanges);
    }

    /// <summary>Extracts structured content for each page, including detected lines, lists, leader rows, and simple tables.</summary>
    public static IReadOnlyList<StructuredPage> ExtractStructuredByPage(byte[] pdf, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        return PdfReadDocument.Load(pdf).ExtractStructuredPages(options);
    }

    /// <summary>Extracts structured content from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredPage> ExtractStructuredByPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        return ExtractStructuredByPageRanges(pdf, (PdfTextLayoutOptions?)null, pageRanges);
    }

    /// <summary>Extracts structured content from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredPage> ExtractStructuredByPageRanges(byte[] pdf, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractStructuredByPageRanges(PdfReadDocument.Load(pdf), options, pageRanges);
    }

    /// <summary>Extracts detected paragraphs grouped by page while preserving paragraph geometry.</summary>
    public static IReadOnlyList<StructuredParagraphPage> ExtractParagraphsByPage(byte[] pdf, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        return PdfReadDocument.Load(pdf).ExtractParagraphsByPage(options);
    }

    /// <summary>Extracts detected paragraphs from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredParagraphPage> ExtractParagraphsByPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        return ExtractParagraphsByPageRanges(pdf, (PdfTextLayoutOptions?)null, pageRanges);
    }

    /// <summary>Extracts detected paragraphs from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredParagraphPage> ExtractParagraphsByPageRanges(byte[] pdf, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractParagraphsByPageRanges(PdfReadDocument.Load(pdf), options, pageRanges);
    }

    /// <summary>Extracts detected headings grouped by page while preserving heading geometry.</summary>
    public static IReadOnlyList<StructuredHeadingPage> ExtractHeadingsByPage(byte[] pdf, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        return PdfReadDocument.Load(pdf).ExtractHeadingsByPage(options);
    }

    /// <summary>Extracts detected headings from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredHeadingPage> ExtractHeadingsByPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        return ExtractHeadingsByPageRanges(pdf, (PdfTextLayoutOptions?)null, pageRanges);
    }

    /// <summary>Extracts detected headings from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredHeadingPage> ExtractHeadingsByPageRanges(byte[] pdf, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractHeadingsByPageRanges(PdfReadDocument.Load(pdf), options, pageRanges);
    }

    /// <summary>Extracts detected list items grouped by page while preserving marker and nesting hints.</summary>
    public static IReadOnlyList<StructuredListItemPage> ExtractListItemsByPage(byte[] pdf, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        return PdfReadDocument.Load(pdf).ExtractListItemsByPage(options);
    }

    /// <summary>Extracts detected list items from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredListItemPage> ExtractListItemsByPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        return ExtractListItemsByPageRanges(pdf, (PdfTextLayoutOptions?)null, pageRanges);
    }

    /// <summary>Extracts detected list items from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredListItemPage> ExtractListItemsByPageRanges(byte[] pdf, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractListItemsByPageRanges(PdfReadDocument.Load(pdf), options, pageRanges);
    }

    /// <summary>Extracts detected tables grouped by page while preserving table geometry.</summary>
    public static IReadOnlyList<StructuredTablePage> ExtractTablesByPage(byte[] pdf, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        return PdfReadDocument.Load(pdf).ExtractTablesByPage(options);
    }

    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredTablePage> ExtractTablesByPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        return ExtractTablesByPageRanges(pdf, (PdfTextLayoutOptions?)null, pageRanges);
    }

    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredTablePage> ExtractTablesByPageRanges(byte[] pdf, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractTablesByPageRanges(PdfReadDocument.Load(pdf), options, pageRanges);
    }

    /// <summary>Extracts plain text from each page using layout options such as column detection and header/footer trimming.</summary>
    public static IReadOnlyList<string> ExtractTextByPage(byte[] pdf, PdfTextLayoutOptions? options) {
        Guard.NotNull(pdf, nameof(pdf));
        if (options is null) {
            return ExtractTextByPage(pdf);
        }

        return ExtractTextByPage(PdfReadDocument.Load(pdf), options);
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<string> ExtractTextByPage(PdfReadDocument document) {
        var pages = new List<string>(document.Pages.Count);
        for (int i = 0; i < document.Pages.Count; i++) {
            pages.Add(document.Pages[i].ExtractText());
        }

        return pages.AsReadOnly();
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<string> ExtractTextByPageRanges(PdfReadDocument document, PdfPageRange[] pageRanges) {
        var selected = ExtractSelectedTextPages(document, pageRanges);
        var pages = new List<string>(selected.Count);
        for (int i = 0; i < selected.Count; i++) {
            pages.Add(selected[i].Text);
        }

        return pages.AsReadOnly();
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<SelectedTextPage> ExtractSelectedTextPages(PdfReadDocument document, PdfPageRange[] pageRanges) {
        return ExtractSelectedTextPages(document, null, pageRanges);
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<SelectedTextPage> ExtractSelectedTextPages(PdfReadDocument document, PdfTextLayoutOptions? options, PdfPageRange[] pageRanges) {
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));

        var pages = new List<SelectedTextPage>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            string text = options is null
                ? document.Pages[pageNumber - 1].ExtractText()
                : document.Pages[pageNumber - 1].ExtractTextWithColumns(options);
            pages.Add(new SelectedTextPage(pageNumber, text));
        }

        return pages.AsReadOnly();
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<StructuredPage> ExtractStructuredByPageRanges(PdfReadDocument document, PdfTextLayoutOptions? options, PdfPageRange[] pageRanges) {
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));

        var pages = new List<StructuredPage>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            pages.Add(document.Pages[pageNumbers[i] - 1].ExtractStructured(options));
        }

        return pages.AsReadOnly();
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<StructuredParagraphPage> ExtractParagraphsByPageRanges(PdfReadDocument document, PdfTextLayoutOptions? options, PdfPageRange[] pageRanges) {
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));

        var pages = new List<StructuredParagraphPage>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            var structuredPage = document.Pages[pageNumber - 1].ExtractStructured(options);
            pages.Add(new StructuredParagraphPage(pageNumber, structuredPage.Paragraphs));
        }

        return pages.AsReadOnly();
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<StructuredHeadingPage> ExtractHeadingsByPageRanges(PdfReadDocument document, PdfTextLayoutOptions? options, PdfPageRange[] pageRanges) {
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));

        var pages = new List<StructuredHeadingPage>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            var structuredPage = document.Pages[pageNumber - 1].ExtractStructured(options);
            pages.Add(new StructuredHeadingPage(pageNumber, structuredPage.Headings));
        }

        return pages.AsReadOnly();
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<StructuredListItemPage> ExtractListItemsByPageRanges(PdfReadDocument document, PdfTextLayoutOptions? options, PdfPageRange[] pageRanges) {
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));

        var pages = new List<StructuredListItemPage>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            var structuredPage = document.Pages[pageNumber - 1].ExtractStructured(options);
            pages.Add(new StructuredListItemPage(pageNumber, structuredPage.ListNodes));
        }

        return pages.AsReadOnly();
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<StructuredTablePage> ExtractTablesByPageRanges(PdfReadDocument document, PdfTextLayoutOptions? options, PdfPageRange[] pageRanges) {
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));

        var pages = new List<StructuredTablePage>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            var structuredPage = document.Pages[pageNumber - 1].ExtractStructured(options);
            pages.Add(new StructuredTablePage(pageNumber, structuredPage.TablesDetailed));
        }

        return pages.AsReadOnly();
    }

    private readonly struct SelectedTextPage {
        internal SelectedTextPage(int pageNumber, string text) {
            PageNumber = pageNumber;
            Text = text;
        }

        internal int PageNumber { get; }
        internal string Text { get; }
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<string> ExtractTextByPage(PdfReadDocument document, PdfTextLayoutOptions options) {
        var pages = new List<string>(document.Pages.Count);
        for (int i = 0; i < document.Pages.Count; i++) {
            pages.Add(document.Pages[i].ExtractTextWithColumns(options));
        }

        return pages.AsReadOnly();
    }

    private static string ExtractTextFromPage(PdfDictionary page, Dictionary<int, PdfIndirectObject> parsedObjects, Dictionary<int, string> rawObjects) {
        var pageText = new StringBuilder();
        var resources = ResolveDict(GetInheritedValue(page, "Resources", parsedObjects), parsedObjects);
        var activeForms = new HashSet<PdfStream>();

        foreach (int contentId in GetContentIds(page, parsedObjects)) {
            if (TryGetContentStreamContent(parsedObjects, rawObjects, contentId, out string content)) {
                pageText.Append(ExtractTextFromContentStream(content, resources, parsedObjects, rawObjects, activeForms));
            }
        }

        return pageText.ToString();
    }

    private static bool TryGetContentStreamContent(Dictionary<int, PdfIndirectObject> parsedObjects, Dictionary<int, string> rawObjects, int contentId, out string content) {
        if (parsedObjects.TryGetValue(contentId, out var parsedObject) &&
            parsedObject.Value is PdfStream stream) {
            byte[] streamBytes = Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, parsedObjects);

            content = PdfEncoding.Latin1GetString(streamBytes);
            return true;
        }

        if (rawObjects.TryGetValue(contentId, out var obj)) {
            var match = StreamRegex.Match(obj);
            if (match.Success) {
                content = match.Groups[1].Value;
                return true;
            }
        }

        content = string.Empty;
        return false;
    }

    /// <summary>Gets document metadata (Title/Author/Subject/Keywords) if present; null when absent.</summary>
    public static (string? Title, string? Author, string? Subject, string? Keywords) GetMetadata(string path) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return GetMetadata(File.ReadAllBytes(path));
    }

    /// <summary>Gets document metadata (Title/Author/Subject/Keywords) from the current position of a readable stream.</summary>
    public static (string? Title, string? Author, string? Subject, string? Keywords) GetMetadata(Stream stream) {
        return GetMetadata(ReadAllBytes(stream));
    }

    /// <summary>Gets document metadata (Title/Author/Subject/Keywords) if present; null when absent.</summary>
    public static (string? Title, string? Author, string? Subject, string? Keywords) GetMetadata(byte[] pdf) {
        var map = BuildObjectMap(pdf, out var trailer);
        var m = InfoRefRegex.Match(trailer);
        if (!m.Success) return (null, null, null, null);
        int infoId = int.Parse(m.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
        if (!map.TryGetValue(infoId, out var obj)) return (null, null, null, null);
        string? title = ExtractStringValue(obj, "/Title");
        string? author = ExtractStringValue(obj, "/Author");
        string? subject = ExtractStringValue(obj, "/Subject");
        string? keywords = ExtractStringValue(obj, "/Keywords");
        return (title, author, subject, keywords);
    }

    private static byte[] ReadAllBytes(Stream stream) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    private static Dictionary<int, string> BuildObjectMap(byte[] pdf, out string trailer) {
        string text = PdfEncoding.Latin1GetString(pdf);
        var dict = new Dictionary<int, string>();
        var matches = ObjRegex.Matches(text);
        for (int i = 0; i < matches.Count; i++) {
            int id = int.Parse(matches[i].Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            int start = matches[i].Index;
            int end = (i + 1 < matches.Count) ? matches[i + 1].Index : text.Length;
            string body = text.Substring(start, end - start);
            // trim header to just 'obj .. endobj'
            int objStart = body.IndexOf("obj", StringComparison.Ordinal);
            int objEnd = body.IndexOf("endobj", StringComparison.Ordinal);
            if (objStart >= 0 && objEnd > objStart) {
                dict[id] = body.Substring(objStart + 3, objEnd - (objStart + 3));
            }
        }
        int trailerIdx = text.LastIndexOf("trailer", StringComparison.OrdinalIgnoreCase);
        trailer = trailerIdx >= 0 ? text.Substring(trailerIdx) : string.Empty;
        PdfSyntax.ThrowIfEncrypted(trailer);
        return dict;
    }

    private static List<List<int>> FindPageContentIds(byte[] pdf) {
        string text = PdfEncoding.Latin1GetString(pdf);
        var ids = new List<List<int>>();
        foreach (Match m in PageObjRegex.Matches(text)) {
            if (m.Groups["single"].Success && int.TryParse(m.Groups["single"].Value, out int singleId)) {
                ids.Add(new List<int> { singleId });
                continue;
            }

            if (!m.Groups["array"].Success) {
                continue;
            }

            var pageIds = new List<int>();
            foreach (Match refMatch in RefRegex.Matches(m.Groups["array"].Value)) {
                if (int.TryParse(refMatch.Groups[1].Value, out int id)) {
                    pageIds.Add(id);
                }
            }

            if (pageIds.Count > 0) {
                ids.Add(pageIds);
            }
        }
        return ids;
    }

    private static string ExtractTextFromContentStream(
        string content,
        PdfDictionary? resources = null,
        Dictionary<int, PdfIndirectObject>? parsedObjects = null,
        Dictionary<int, string>? rawObjects = null,
        HashSet<PdfStream>? activeForms = null) {
        var sb = new StringBuilder();
        bool inText = false;
        bool pendingSpace = false;
        bool hasTextInCurrentTextObject = false;
        double currentFontSize = 12;
        double currentHorizontalScale = 1.0;
        var args = new List<object>(8);
        int i = 0;
        int n = content.Length;
        while (i < n) {
            SkipWs();
            if (i >= n) break;

            char c = content[i];
            if (c == '%') {
                while (i < n && content[i] != '\n' && content[i] != '\r') i++;
                continue;
            }

            if (c == '/') { args.Add(ReadName()); continue; }
            if (c == '(') { args.Add(ReadLiteralString()); continue; }
            if (c == '<') {
                if (i + 1 < n && content[i + 1] == '<') { i += 2; continue; }
                args.Add(ReadHexString());
                continue;
            }
            if (c == '[') { args.Add(ReadArray()); continue; }
            if (c == ']' || c == '>') { i++; continue; }
            if (IsNumberStart(c)) { args.Add(ReadNumber()); continue; }

            string op = ReadOperator();
            if (op.Length == 0) {
                i++;
                continue;
            }

            switch (op) {
                case "BT":
                    inText = true;
                    pendingSpace = false;
                    hasTextInCurrentTextObject = false;
                    args.Clear();
                    break;
                case "ET":
                    inText = false;
                    pendingSpace = false;
                    hasTextInCurrentTextObject = false;
                    args.Clear();
                    break;
                case "T*":
                    if (inText) {
                        sb.AppendLine();
                        pendingSpace = false;
                    }
                    args.Clear();
                    break;
                case "Tf":
                    if (args.Count >= 2) {
                        currentFontSize = ToDouble(args[args.Count - 1]);
                    }
                    args.Clear();
                    break;
                case "Tz":
                    if (args.Count >= 1) {
                        currentHorizontalScale = ToDouble(args[args.Count - 1]) / 100.0;
                    }
                    args.Clear();
                    break;
                case "Td":
                case "TD":
                    if (inText && args.Count >= 2) {
                        double advanceX = ToDouble(args[args.Count - 2]);
                        double advanceY = ToDouble(args[args.Count - 1]);
                        if (Math.Abs(advanceY) > 0.1 && hasTextInCurrentTextObject) {
                            AppendLineBreak();
                        } else if (advanceX > 0.1) {
                            pendingSpace = true;
                        }
                    }
                    args.Clear();
                    break;
                case "Tj":
                    if (inText && args.Count >= 1) {
                        AppendTextRun(ToText(args[args.Count - 1]));
                    }
                    args.Clear();
                    break;
                case "TJ":
                    if (inText && args.Count >= 1) {
                        AppendTextArray(args[args.Count - 1]);
                    }
                    args.Clear();
                    break;
                case "'":
                    if (inText && args.Count >= 1) {
                        AppendLineBreak();
                        AppendTextRun(ToText(args[args.Count - 1]));
                    }
                    args.Clear();
                    break;
                case "\"":
                    if (inText && args.Count >= 3) {
                        AppendLineBreak();
                        AppendTextRun(ToText(args[args.Count - 1]));
                    }
                    args.Clear();
                    break;
                case "Do":
                    if (resources is not null && parsedObjects is not null && rawObjects is not null && args.Count >= 1) {
                        string formText = ExtractInvokedFormText(ToName(args[args.Count - 1]), resources, parsedObjects, rawObjects, activeForms ?? new HashSet<PdfStream>());
                        if (!string.IsNullOrEmpty(formText)) {
                            AppendTextRun(formText);
                        }
                    }
                    args.Clear();
                    break;
                default:
                    args.Clear();
                    break;
            }
        }

        return sb.ToString();

        void AppendTextRun(string value) {
            if (string.IsNullOrEmpty(value)) return;
            if (pendingSpace &&
                sb.Length > 0 &&
                !char.IsWhiteSpace(sb[sb.Length - 1]) &&
                !char.IsWhiteSpace(value[0])) {
                sb.Append(' ');
            }
            sb.Append(value);
            if (inText) {
                hasTextInCurrentTextObject = true;
            }
            pendingSpace = false;
        }

        void RequestSpace() {
            pendingSpace = true;
        }

        void AppendLineBreak() {
            if (sb.Length > 0) {
                sb.AppendLine();
            }
            pendingSpace = false;
        }

        void AppendTextArray(object arrayObject) {
            if (arrayObject is not List<object> list) {
                return;
            }

            foreach (var item in list) {
                if (item is string text) {
                    AppendTextRun(text);
                } else if (item is double adjustment &&
                           (-adjustment / 1000.0 * currentFontSize * currentHorizontalScale) > Math.Max(1.5, currentFontSize * 0.24)) {
                    RequestSpace();
                }
            }
        }

        void SkipWs() {
            while (i < n && char.IsWhiteSpace(content[i])) i++;
        }

        double ReadNumber() {
            int start = i;
            i++;
            while (i < n) {
                char ch = content[i];
                if (!(char.IsDigit(ch) || ch == '.' || ch == 'E' || ch == 'e' || ch == '-' || ch == '+')) break;
                i++;
            }

            var s = content.Substring(start, i - start);
            if (!double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v)) v = 0;
            return v;
        }

        string ReadName() {
            i++;
            int start = i;
            while (i < n) {
                char ch = content[i];
                if (char.IsWhiteSpace(ch) || ch == '%' || ch == '/' || ch == '[' || ch == ']' || ch == '(' || ch == ')' || ch == '<' || ch == '>') break;
                i++;
            }
            return PdfSyntax.DecodeName(content.Substring(start, i - start));
        }

        string ReadLiteralString() {
            i++;
            int depth = 1;
            bool escaped = false;
            var raw = new StringBuilder();
            while (i < n && depth > 0) {
                char ch = content[i++];
                if (escaped) {
                    raw.Append(ch);
                    escaped = false;
                    continue;
                }

                if (ch == '\\') {
                    raw.Append(ch);
                    escaped = true;
                    continue;
                }

                if (ch == '(') {
                    depth++;
                    raw.Append(ch);
                    continue;
                }

                if (ch == ')') {
                    depth--;
                    if (depth > 0) {
                        raw.Append(ch);
                    }
                    continue;
                }

                raw.Append(ch);
            }

            return UnescapePdfLiteral(raw.ToString());
        }

        string ReadHexString() {
            i++;
            var raw = new StringBuilder();
            while (i < n && content[i] != '>') {
                raw.Append(content[i]);
                i++;
            }

            if (i < n && content[i] == '>') {
                i++;
            }

            return DecodeHexPdfString(raw.ToString());
        }

        List<object> ReadArray() {
            var list = new List<object>();
            i++;
            while (i < n) {
                SkipWs();
                if (i >= n) break;
                char ch = content[i];
                if (ch == ']') {
                    i++;
                    break;
                }

                if (ch == '(') {
                    list.Add(ReadLiteralString());
                    continue;
                }

                if (ch == '<') {
                    if (i + 1 < n && content[i + 1] == '<') {
                        i += 2;
                        continue;
                    }
                    list.Add(ReadHexString());
                    continue;
                }

                if (IsNumberStart(ch)) {
                    list.Add(ReadNumber());
                    continue;
                }

                if (ch == '/') {
                    list.Add(ReadName());
                    continue;
                }

                if (ch == '[') {
                    i++;
                    continue;
                }

                ReadOperator();
            }

            return list;
        }

        string ReadOperator() {
            int start = i;
            char ch = content[i++];
            if (ch == '\'' || ch == '"') return ch.ToString();
            while (i < n) {
                char current = content[i];
                if (char.IsWhiteSpace(current) || current == '%' || current == '(' || current == '[' || current == '/' || current == '<' || current == '>') break;
                i++;
            }
            return content.Substring(start, i - start);
        }

        static bool IsNumberStart(char ch) => ch == '+' || ch == '-' || ch == '.' || char.IsDigit(ch);
        static double ToDouble(object o) => o is double d ? d : 0.0;
        static string ToText(object o) => o as string ?? string.Empty;
        static string ToName(object o) => o as string ?? string.Empty;
    }

    private static string ExtractInvokedFormText(
        string formName,
        PdfDictionary resources,
        Dictionary<int, PdfIndirectObject> parsedObjects,
        Dictionary<int, string> rawObjects,
        HashSet<PdfStream> activeForms) {
        if (!TryGetFormStream(resources, formName, parsedObjects, out var formStream)) {
            return string.Empty;
        }

        if (!activeForms.Add(formStream)) {
            return string.Empty;
        }

        try {
            string content = DecodeStreamContent(formStream, parsedObjects);
            var formResources = ResolveDict(formStream.Dictionary.Items.TryGetValue("Resources", out var resObj) ? resObj : null, parsedObjects) ?? resources;
            return ExtractTextFromContentStream(content, formResources, parsedObjects, rawObjects, activeForms);
        } finally {
            activeForms.Remove(formStream);
        }
    }

    private static bool TryGetFormStream(
        PdfDictionary resources,
        string name,
        Dictionary<int, PdfIndirectObject> parsedObjects,
        out PdfStream formStream) {
        formStream = null!;

        if (!resources.Items.TryGetValue("XObject", out var xObjectObj)) {
            return false;
        }

        var xObjectDict = ResolveDict(xObjectObj, parsedObjects);
        if (xObjectDict is null || !xObjectDict.Items.TryGetValue(name, out var formObj)) {
            return false;
        }

        if (formObj is PdfReference formRef &&
            PdfObjectLookup.TryGet(parsedObjects, formRef, out var indirectForm) &&
            indirectForm.Value is PdfStream referencedStream &&
            string.Equals(referencedStream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) {
            formStream = referencedStream;
            return true;
        }

        if (formObj is PdfStream directStream &&
            string.Equals(directStream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) {
            formStream = directStream;
            return true;
        }

        return false;
    }

    private static string DecodeStreamContent(PdfStream stream, Dictionary<int, PdfIndirectObject>? parsedObjects = null) {
        byte[] bytes = Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, parsedObjects);
        return PdfEncoding.Latin1GetString(bytes);
    }

    internal static string UnescapePdfLiteral(string s) {
        var sb = new StringBuilder();
        for (int i = 0; i < s.Length; i++) {
            char c = s[i];
            if (c == '\\' && i + 1 < s.Length) {
                char n = s[++i];
                if (n >= '0' && n <= '7') {
                    int value = n - '0';
                    int digits = 1;
                    while (digits < 3 && i + 1 < s.Length && s[i + 1] >= '0' && s[i + 1] <= '7') {
                        value = (value * 8) + (s[++i] - '0');
                        digits++;
                    }
                    sb.Append((char)value);
                    continue;
                }

                if (n == '\r') {
                    if (i + 1 < s.Length && s[i + 1] == '\n') {
                        i++;
                    }
                    continue;
                }

                if (n == '\n') {
                    continue;
                }

                sb.Append(n switch {
                    'n' => '\n',
                    'r' => '\r',
                    't' => '\t',
                    'b' => '\b',
                    'f' => '\f',
                    '(' => '(',
                    ')' => ')',
                    '\\' => '\\',
                    _ => n
                });
            } else sb.Append(c);
        }
        return sb.ToString();
    }

    internal static string DecodeHexPdfString(string s) {
        if (string.IsNullOrWhiteSpace(s)) return string.Empty;

        var hex = new StringBuilder(s.Length);
        for (int i = 0; i < s.Length; i++) {
            char ch = s[i];
            if (!char.IsWhiteSpace(ch)) hex.Append(ch);
        }

        if (hex.Length % 2 == 1) hex.Append('0');

        var bytes = new byte[hex.Length / 2];
        for (int i = 0; i < bytes.Length; i++) {
            int hi = HexNibble(hex[i * 2]);
            int lo = HexNibble(hex[i * 2 + 1]);
            bytes[i] = (byte)((hi << 4) | lo);
        }

        return PdfWinAnsiEncoding.Decode(bytes);

        static int HexNibble(char c) {
            if (c >= '0' && c <= '9') return c - '0';
            if (c >= 'a' && c <= 'f') return 10 + (c - 'a');
            if (c >= 'A' && c <= 'F') return 10 + (c - 'A');
            throw new FormatException($"Invalid hex character '{c}'.");
        }
    }

    private static string? ExtractStringValue(string obj, string key) {
        int idx = obj.IndexOf(key, StringComparison.Ordinal);
        if (idx < 0) return null;
        int valueStart = idx + key.Length;
        while (valueStart < obj.Length && char.IsWhiteSpace(obj[valueStart])) {
            valueStart++;
        }

        if (valueStart >= obj.Length) {
            return null;
        }

        if (obj[valueStart] == '(') {
            int close = FindCloseParen(obj, valueStart);
            if (close < 0) return null;
            string raw = obj.Substring(valueStart + 1, close - valueStart - 1);
            return PdfTextString.DecodeLiteral(raw);
        }

        if (obj[valueStart] == '<' && (valueStart + 1 >= obj.Length || obj[valueStart + 1] != '<')) {
            int close = obj.IndexOf('>', valueStart + 1);
            if (close < 0) return null;
            string raw = obj.Substring(valueStart + 1, close - valueStart - 1);
            return DecodeMetadataHexString(raw);
        }

        return null;
    }

    private static int FindCloseParen(string s, int start) {
        int depth = 0;
        for (int i = start; i < s.Length; i++) {
            char c = s[i];
            if (c == '\\') { i++; continue; }
            if (c == '(') depth++;
            else if (c == ')') { depth--; if (depth == 0) return i; }
        }
        return -1;
    }

    private static bool TryReadTextAdvance(string line, out double advanceX, out double advanceY) {
        advanceX = 0;
        advanceY = 0;

        if (!line.EndsWith(" Td", StringComparison.Ordinal) && line != "Td") {
            return false;
        }

        var parts = line.Split(SpaceSplitChars, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 3 || !string.Equals(parts[parts.Length - 1], "Td", StringComparison.Ordinal)) {
            return false;
        }

        return double.TryParse(parts[parts.Length - 3], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out advanceX) &&
               double.TryParse(parts[parts.Length - 2], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out advanceY);
    }

    private static bool TryReadFontSize(string line, out double fontSize) {
        fontSize = 0;
        var parts = line.Split(SpaceSplitChars, StringSplitOptions.RemoveEmptyEntries);
        return parts.Length >= 3 &&
               string.Equals(parts[parts.Length - 1], "Tf", StringComparison.Ordinal) &&
               double.TryParse(parts[parts.Length - 2], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out fontSize);
    }

    private static bool TryReadHorizontalScale(string line, out double horizontalScale) {
        horizontalScale = 0;
        var parts = line.Split(SpaceSplitChars, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 2 ||
            !string.Equals(parts[parts.Length - 1], "Tz", StringComparison.Ordinal) ||
            !double.TryParse(parts[parts.Length - 2], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double percent)) {
            return false;
        }

        horizontalScale = percent / 100.0;
        return true;
    }

    private static bool TryAppendTextArray(string line, double fontSize, double horizontalScale, Action<string> appendText, Action requestSpace) {
        int operatorIndex = line.LastIndexOf("TJ", StringComparison.Ordinal);
        if (operatorIndex < 0) {
            return false;
        }

        int arrayStart = line.IndexOf('[');
        int arrayEnd = line.LastIndexOf(']');
        if (arrayStart < 0 || arrayEnd < arrayStart || arrayEnd > operatorIndex) {
            return false;
        }

        string items = line.Substring(arrayStart + 1, arrayEnd - arrayStart - 1);
        int i = 0;
        while (i < items.Length) {
            while (i < items.Length && char.IsWhiteSpace(items[i])) {
                i++;
            }

            if (i >= items.Length) {
                break;
            }

            char ch = items[i];
            if (ch == '(') {
                appendText(ReadLiteral(items, ref i));
                continue;
            }

            if (ch == '<') {
                appendText(ReadHex(items, ref i));
                continue;
            }

            if (IsNumberStart(ch)) {
                if (TryReadNumber(items, ref i, out double adjustment) &&
                    ToVisualAdvance(adjustment, fontSize, horizontalScale) > Math.Max(1.5, fontSize * 0.24)) {
                    requestSpace();
                }
                continue;
            }

            i++;
        }

        return true;

        static string ReadLiteral(string s, ref int index) {
            index++; // skip (
            int depth = 1;
            bool escaped = false;
            var raw = new StringBuilder();
            while (index < s.Length && depth > 0) {
                char current = s[index++];
                if (escaped) {
                    raw.Append(current);
                    escaped = false;
                    continue;
                }

                if (current == '\\') {
                    raw.Append(current);
                    escaped = true;
                    continue;
                }

                if (current == '(') {
                    depth++;
                    raw.Append(current);
                    continue;
                }

                if (current == ')') {
                    depth--;
                    if (depth > 0) {
                        raw.Append(current);
                    }
                    continue;
                }

                raw.Append(current);
            }

            return UnescapePdfLiteral(raw.ToString());
        }

        static string ReadHex(string s, ref int index) {
            index++; // skip <
            var raw = new StringBuilder();
            while (index < s.Length && s[index] != '>') {
                raw.Append(s[index]);
                index++;
            }

            if (index < s.Length && s[index] == '>') {
                index++;
            }

            return DecodeHexPdfString(raw.ToString());
        }

        static bool TryReadNumber(string s, ref int index, out double value) {
            int start = index;
            if (s[index] == '+' || s[index] == '-') {
                index++;
            }

            while (index < s.Length && (char.IsDigit(s[index]) || s[index] == '.')) {
                index++;
            }

#if NET8_0_OR_GREATER
            return double.TryParse(s.AsSpan(start, index - start), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out value);
#else
            return double.TryParse(s.Substring(start, index - start), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out value);
#endif
        }

        static bool IsNumberStart(char ch) => ch == '+' || ch == '-' || ch == '.' || char.IsDigit(ch);
        static double ToVisualAdvance(double tjAdjustment, double currentFontSize, double currentHorizontalScale) => -tjAdjustment / 1000.0 * currentFontSize * currentHorizontalScale;
    }

    private static bool TryAppendNextLineShowText(string line, Action<string> appendText, Action requestSpace) {
        Match literalQuote = QuoteLiteralRegex.Match(line);
        if (literalQuote.Success) {
            requestSpace();
            appendText(UnescapePdfLiteral(literalQuote.Groups["txt"].Value));
            return true;
        }

        Match hexQuote = QuoteHexRegex.Match(line);
        if (hexQuote.Success) {
            requestSpace();
            appendText(DecodeHexPdfString(hexQuote.Groups["txt"].Value));
            return true;
        }

        Match literalDoubleQuote = DoubleQuoteLiteralRegex.Match(line);
        if (literalDoubleQuote.Success) {
            requestSpace();
            appendText(UnescapePdfLiteral(literalDoubleQuote.Groups["txt"].Value));
            return true;
        }

        Match hexDoubleQuote = DoubleQuoteHexRegex.Match(line);
        if (hexDoubleQuote.Success) {
            requestSpace();
            appendText(DecodeHexPdfString(hexDoubleQuote.Groups["txt"].Value));
            return true;
        }

        return false;
    }

    private static string DecodeMetadataHexString(string raw) {
        return PdfTextString.DecodeHex(raw);
    }

    private static List<PdfDictionary> CollectPages(Dictionary<int, PdfIndirectObject> objects, string? trailerRaw) {
        var result = new List<PdfDictionary>();
        PdfDictionary? catalog = PdfSyntax.FindCatalog(objects, trailerRaw);
        if (catalog is not null &&
            ResolveDict(catalog.Items.TryGetValue("Pages", out var pagesObj) ? pagesObj : null, objects) is PdfDictionary pagesRoot) {
            TraversePagesNode(pagesRoot, objects, result, new HashSet<int>());
        }

        if (result.Count > 0) {
            return result;
        }

        foreach (var kv in objects.OrderBy(k => k.Key)) {
            if (kv.Value.Value is PdfDictionary dict && IsLeafPage(dict, objects)) {
                result.Add(dict);
            }
        }

        return result;
    }

    private static void TraversePagesNode(
        PdfDictionary node,
        Dictionary<int, PdfIndirectObject> objects,
        List<PdfDictionary> result,
        HashSet<int> visited) {
        int objectNumber = FindObjectNumberFor(node, objects);
        if (objectNumber > 0 && !visited.Add(objectNumber)) {
            return;
        }

        string? type = node.Get<PdfName>("Type")?.Name;
        if (type == "Page" || (type is null && IsLeafPage(node, objects))) {
            result.Add(node);
            return;
        }

        var kids = ResolveArray(node.Items.TryGetValue("Kids", out var kidsObj) ? kidsObj : null, objects);
        if (kids is null) {
            return;
        }

        foreach (var kid in kids.Items) {
            var child = ResolveDict(kid, objects);
            if (child is not null) {
                TraversePagesNode(child, objects, result, visited);
            }
        }
    }

    private static bool IsLeafPage(PdfDictionary page, Dictionary<int, PdfIndirectObject> objects) {
        if (ResolveArray(page.Items.TryGetValue("Kids", out var kidsObj) ? kidsObj : null, objects) is not null) {
            return false;
        }

        if (!page.Items.ContainsKey("Contents")) {
            return false;
        }

        string? type = page.Get<PdfName>("Type")?.Name;
        if (type == "Page") {
            return true;
        }

        return type is null &&
               (page.Items.ContainsKey("Resources") || GetInheritedValue(page, "Resources", objects) is not null) &&
               (page.Items.ContainsKey("MediaBox") ||
                page.Items.ContainsKey("CropBox") ||
                GetInheritedValue(page, "MediaBox", objects) is not null ||
                GetInheritedValue(page, "CropBox", objects) is not null);
    }

    private static List<int> GetContentIds(PdfDictionary page, Dictionary<int, PdfIndirectObject> objects) {
        var ids = new List<int>();
        if (!page.Items.TryGetValue("Contents", out var contents)) {
            return ids;
        }

        if (contents is PdfReference reference) {
            if (PdfObjectLookup.TryGet(objects, reference, out var indirect)) {
                if (indirect.Value is PdfArray referencedArray) {
                    AppendContentIds(referencedArray, ids, objects);
                } else if (indirect.Value is PdfStream) {
                    ids.Add(reference.ObjectNumber);
                }
            }
            return ids;
        }

        if (contents is PdfArray arr) {
            AppendContentIds(arr, ids, objects);
        }

        return ids;
    }

    private static PdfObject? GetInheritedValue(PdfDictionary start, string key, Dictionary<int, PdfIndirectObject> objects) {
        PdfDictionary? current = start;
        int guard = 0;
        while (current is not null && guard++ < 100) {
            if (current.Items.TryGetValue(key, out var value)) {
                return value;
            }

            if (!current.Items.TryGetValue("Parent", out var parentObj)) {
                break;
            }

            current = ResolveDict(parentObj, objects);
        }

        return null;
    }

    private static PdfDictionary? ResolveDict(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        if (obj is PdfDictionary dict) {
            return dict;
        }

        if (obj is PdfReference reference &&
            PdfObjectLookup.TryGet(objects, reference, out var indirect) &&
            indirect.Value is PdfDictionary referencedDict) {
            return referencedDict;
        }

        return null;
    }

    private static PdfArray? ResolveArray(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        if (obj is PdfArray arr) {
            return arr;
        }

        if (obj is PdfReference reference &&
            PdfObjectLookup.TryGet(objects, reference, out var indirect) &&
            indirect.Value is PdfArray referencedArray) {
            return referencedArray;
        }

        return null;
    }

    private static void AppendContentIds(PdfArray contentArray, List<int> ids, Dictionary<int, PdfIndirectObject> objects) {
        foreach (var item in contentArray.Items) {
            if (item is PdfReference itemReference &&
                PdfObjectLookup.TryGet(objects, itemReference, out var indirect) &&
                indirect.Value is PdfStream) {
                ids.Add(itemReference.ObjectNumber);
            }
        }
    }

    private static int FindObjectNumberFor(PdfDictionary dict, Dictionary<int, PdfIndirectObject> objects) {
        foreach (var kv in objects) {
            if (ReferenceEquals(kv.Value.Value, dict)) {
                return kv.Key;
            }
        }

        return 0;
    }
}
