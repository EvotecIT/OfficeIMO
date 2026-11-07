using OfficeIMO.Excel;
using OfficeIMO.Reader.FormatInternals;

namespace OfficeIMO.Reader.Excel;

internal static partial class ExcelReaderAdapter {
    private static ExcelDocument LoadForProjection(string path, ReaderOptions options, CancellationToken token,
        out IReadOnlyList<string>? warnings) {
        warnings = null;
        if (options.DetectionMode != ReaderDetectionMode.PreferContent &&
            ExcelFormatCatalog.TryGetByExtension(path, out _)) return Load(path, options, token);

        using var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return HasSpreadsheetXmlRoot(input, options, token)
            ? ImportSpreadsheetXml(input, options, token, out warnings)
            : Load(path, options, token);
    }

    private static ExcelDocument LoadForProjection(Stream input, string? sourceName, ReaderOptions options,
        CancellationToken token, out IReadOnlyList<string>? warnings) {
        warnings = null;
        if ((options.DetectionMode != ReaderDetectionMode.PreferContent &&
             ExcelFormatCatalog.TryGetByExtension(sourceName, out _)) ||
            !HasSpreadsheetXmlRoot(input, options, token)) return Load(input, sourceName, options, token);

        return ImportSpreadsheetXml(input, options, token, out warnings);
    }

    private static ExcelDocument ImportSpreadsheetXml(Stream input, ReaderOptions options, CancellationToken token,
        out IReadOnlyList<string>? warnings) {
        var settings = new ExcelSpreadsheetXmlOptions { CancellationToken = token };
        if (options.MaxInputBytes is long limit)
            settings.MaxInputBytes = (int)Math.Min(settings.MaxInputBytes, limit);
        var result = NativeReaderAdmission.Load(() => ExcelDocument.ImportSpreadsheetXml(input, settings));
        warnings = result.Report.Diagnostics.Count == 0
            ? null : result.Report.Diagnostics.Select(diagnostic => $"{diagnostic.Code}: {diagnostic.Message}").ToArray();
        return result.Value;
    }

    private static bool HasSpreadsheetXmlRoot(Stream input, ReaderOptions options, CancellationToken token) =>
        DocumentReaderEngine.HasXmlRoot(input, "Workbook", "urn:schemas-microsoft-com:office:spreadsheet",
            options.DetectionMaxProbeBytes, token);
}
