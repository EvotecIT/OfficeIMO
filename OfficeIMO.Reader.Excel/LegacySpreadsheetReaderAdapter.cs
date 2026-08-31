using OfficeIMO.Excel.Legacy;

namespace OfficeIMO.Reader.Excel;

internal static class LegacySpreadsheetReaderAdapter {
    internal static LegacySpreadsheetImportOptions? Clone(LegacySpreadsheetImportOptions? source) {
        if (source == null) return null;
        return new LegacySpreadsheetImportOptions {
            Limits = (source.Limits ?? new OfficeLegacyImportLimits()).Clone(),
            FormatHint = source.FormatHint,
            SourceName = source.SourceName,
            RequireStructured = source.RequireStructured
        };
    }

    internal static OfficeDocumentReadResult ReadDocument(string path, ReaderOptions readerOptions, ReaderExcelOptions options,
        LegacySpreadsheetImportOptions? importOptions, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import(path, Prepare(importOptions, path, readerOptions), cancellationToken);
        return ExcelReaderAdapter.Project(imported.Document, path, readerOptions, options, cancellationToken, BuildWarnings(imported));
    }

    internal static OfficeDocumentReadResult ReadDocument(Stream stream, string? sourceName, ReaderOptions readerOptions, ReaderExcelOptions options,
        LegacySpreadsheetImportOptions? importOptions, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        string logicalName = string.IsNullOrWhiteSpace(sourceName) ? "legacy-workbook" : sourceName!;
        using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import(stream, Prepare(importOptions, sourceName, readerOptions), cancellationToken);
        return ExcelReaderAdapter.Project(imported.Document, logicalName, readerOptions, options, cancellationToken, BuildWarnings(imported));
    }

    internal static bool Probe(Stream stream, string? sourceName, ReaderOptions readerOptions,
        LegacySpreadsheetImportOptions? importOptions, CancellationToken cancellationToken) {
        if (!stream.CanSeek) return false;
        long position = stream.Position;
        try {
            LegacySpreadsheetImportOptions configured = Prepare(importOptions, sourceName, readerOptions);
            long remaining = stream.Length - position;
            if (remaining < 0 || remaining > configured.Limits.MaxInputBytes || remaining > int.MaxValue) return false;
            var data = new byte[(int)remaining];
            int total = 0;
            while (total < data.Length) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = stream.Read(data, total, data.Length - total);
                if (read == 0) break;
                total += read;
            }
            if (total != data.Length) return false;
            LegacySpreadsheetImporter.Detect(data, configured, cancellationToken);
            return true;
        } catch (OperationCanceledException) {
            throw;
        } catch (InvalidDataException) {
            return false;
        } catch (IOException) {
            return false;
        } catch (NotSupportedException) {
            return false;
        } finally {
            stream.Position = position;
        }
    }

    private static LegacySpreadsheetImportOptions Prepare(LegacySpreadsheetImportOptions? source, string? sourceName, ReaderOptions readerOptions) {
        OfficeLegacyImportLimits limits = (source?.Limits ?? new OfficeLegacyImportLimits()).Clone();
        if (readerOptions.MaxInputBytes.HasValue) {
            limits.MaxInputBytes = (int)Math.Min(limits.MaxInputBytes, Math.Min(int.MaxValue, readerOptions.MaxInputBytes.Value));
        }
        return new LegacySpreadsheetImportOptions {
            Limits = limits,
            FormatHint = source?.FormatHint,
            SourceName = string.IsNullOrWhiteSpace(source?.SourceName) ? sourceName : source!.SourceName,
            RequireStructured = source?.RequireStructured ?? false
        };
    }

    private static IReadOnlyList<string> BuildWarnings(LegacySpreadsheetImportResult imported) {
        var warnings = new List<string> { $"Legacy import quality: {imported.Report.Quality}; profile: {imported.Detection.ProfileId}." };
        warnings.AddRange(imported.Report.Findings.Select(static finding => finding.Code + ": " + finding.Message));
        return warnings;
    }
}
