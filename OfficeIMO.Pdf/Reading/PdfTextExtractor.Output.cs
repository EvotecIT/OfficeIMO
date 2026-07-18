using System.Text.RegularExpressions;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Pdf;

internal static partial class PdfTextExtractor {
    private static List<string> WriteTextPages(string baseName, string fullOutputDirectory, IReadOnlyList<string> pages) {
        string safeBaseName = GetSafeBaseName(baseName, "page");
    
        var paths = new List<string>(pages.Count);
        for (int i = 0; i < pages.Count; i++) {
            string outputPath = Path.Combine(fullOutputDirectory, safeBaseName + "-page-" + (i + 1).ToString("0000", System.Globalization.CultureInfo.InvariantCulture) + ".txt");
            OfficeFileCommit.WriteAllBytes(outputPath, Encoding.UTF8.GetBytes(pages[i]));
            paths.Add(outputPath);
        }
    
        return paths;
    }
    
    private static List<string> WriteMarkdownPages(string baseName, string fullOutputDirectory, IReadOnlyList<string> pages) {
        string safeBaseName = GetSafeBaseName(baseName, "page");
    
        var paths = new List<string>(pages.Count);
        for (int i = 0; i < pages.Count; i++) {
            string outputPath = Path.Combine(fullOutputDirectory, safeBaseName + "-page-" + (i + 1).ToString("0000", System.Globalization.CultureInfo.InvariantCulture) + ".md");
            OfficeFileCommit.WriteAllBytes(outputPath, new UTF8Encoding(false).GetBytes(pages[i]));
            paths.Add(outputPath);
        }
    
        return paths;
    }
    
    private static List<string> WriteMarkdownPages(string baseName, string fullOutputDirectory, IReadOnlyList<SelectedTextPage> pages) {
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
                ".md");
            OfficeFileCommit.WriteAllBytes(outputPath, new UTF8Encoding(false).GetBytes(pages[i].Text));
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
            OfficeFileCommit.WriteAllBytes(outputPath, Encoding.UTF8.GetBytes(pages[i].Text));
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
    
                OfficeFileCommit.WriteAllBytes(outputPath, Encoding.UTF8.GetBytes(BuildCsv(page.Tables[tableIndex])));
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
        OfficeStreamWriter.WriteAllBytes(outputStream, bytes);
    }
    
    private static void WriteTextOutput(string outputPath, string text) {
        string fullPath = ValidateOutputPath(outputPath);
        var directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }
    
        OfficeFileCommit.WriteAllBytes(fullPath, new UTF8Encoding(false).GetBytes(text));
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
    
    private static byte[] ReadAllBytes(Stream stream) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
    
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }
    
}
