using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Reader.FormatInternals;

namespace OfficeIMO.Reader.Excel;

internal static class ExcelReaderAdapter {
    internal static ReaderExcelOptions Clone(ReaderExcelOptions? source) => new ReaderExcelOptions {
        SheetName = source?.SheetName,
        A1Range = source?.A1Range,
        HeadersInFirstRow = source?.HeadersInFirstRow ?? true,
        ChunkRows = Math.Max(1, source?.ChunkRows ?? 200),
        ReadOptions = source?.ReadOptions
    };

    internal static OfficeDocumentReadResult ReadDocument(
        string path,
        ReaderOptions readerOptions,
        ReaderExcelOptions options,
        CancellationToken cancellationToken) {
        using ExcelDocument document = Load(path, readerOptions);
        using ExcelDocumentReader reader = document.CreateReader(options.ReadOptions);
        ReaderChunk[] chunks = Extract(reader, path, readerOptions, options, BuildLegacyWarnings(document), cancellationToken).ToArray();
        IReadOnlyList<OfficeDocumentAsset> assets = OpenXmlImageAssetCollector.CollectExcel(
            document.OpenXmlDocument, path, readerOptions, options.SheetName, options.A1Range, cancellationToken);
        OfficeDocumentReadResult result = DocumentReaderEngine.CreateDocumentResult(
            chunks,
            ReaderInputKind.Excel,
            capabilities: new[] { OfficeDocumentReaderBuilderExcelExtensions.HandlerId },
            assets: assets);
        return ExcelRichMapping.Apply(document.CreateInspectionSnapshot(), readerOptions, options, result);
    }

    internal static OfficeDocumentReadResult ReadDocument(
        Stream stream,
        string? sourceName,
        ReaderOptions readerOptions,
        ReaderExcelOptions options,
        CancellationToken cancellationToken) {
        using ExcelDocument document = Load(stream, sourceName, readerOptions);
        using ExcelDocumentReader reader = document.CreateReader(options.ReadOptions);
        string logicalName = string.IsNullOrWhiteSpace(sourceName) ? "workbook.xlsx" : sourceName!;
        ReaderChunk[] chunks = Extract(reader, logicalName, readerOptions, options, BuildLegacyWarnings(document), cancellationToken).ToArray();
        IReadOnlyList<OfficeDocumentAsset> assets = OpenXmlImageAssetCollector.CollectExcel(
            document.OpenXmlDocument, logicalName, readerOptions, options.SheetName, options.A1Range, cancellationToken);
        OfficeDocumentReadResult result = DocumentReaderEngine.CreateDocumentResult(
            chunks,
            ReaderInputKind.Excel,
            capabilities: new[] { OfficeDocumentReaderBuilderExcelExtensions.HandlerId },
            assets: assets);
        return ExcelRichMapping.Apply(document.CreateInspectionSnapshot(), readerOptions, options, result);
    }

    internal static bool ProbeEncryptedOpenXml(
        Stream stream, string? sourceName, ReaderOptions options, CancellationToken cancellationToken) {
        if (string.IsNullOrEmpty(options.OpenPassword) || !stream.CanSeek) return false;
        long position = stream.Position;
        try {
            cancellationToken.ThrowIfCancellationRequested();
            using ExcelDocument document = Load(stream, sourceName, options);
            cancellationToken.ThrowIfCancellationRequested();
            return document.OpenXmlDocument.GetAllParts().Any(static part =>
                string.Equals(part.Uri.OriginalString, "/xl/workbook.xml",
                    StringComparison.OrdinalIgnoreCase));
        } catch (OperationCanceledException) {
            throw;
        } catch {
            return false;
        } finally {
            stream.Position = position;
        }
    }

    private static ExcelDocument Load(string path, ReaderOptions options) {
        string extension = Path.GetExtension(path);
        if (string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(options.OpenPassword)) {
            return ExcelDocument.LoadLegacyXls(path, new LegacyXlsImportOptions { Password = options.OpenPassword, ReportUnsupportedContent = true });
        }
        var loadOptions = new ExcelLoadOptions {
            AccessMode = DocumentAccessMode.ReadOnly,
            OpenSettings = options.OpenXmlMaxCharactersInPart.HasValue
                ? new OpenSettings { MaxCharactersInPart = options.OpenXmlMaxCharactersInPart.Value }
                : null
        };
        try {
            return ExcelDocument.Load(path, loadOptions);
        } catch (Exception exception) when (!string.IsNullOrEmpty(options.OpenPassword) && exception is InvalidDataException or IOException) {
            return ExcelDocument.LoadEncrypted(path, options.OpenPassword!, loadOptions);
        }
    }

    private static ExcelDocument Load(Stream stream, string? sourceName, ReaderOptions options) {
        string extension = Path.GetExtension(sourceName ?? string.Empty);
        if (string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(options.OpenPassword)) {
            return ExcelDocument.LoadLegacyXls(stream, new LegacyXlsImportOptions { Password = options.OpenPassword, ReportUnsupportedContent = true });
        }
        var loadOptions = new ExcelLoadOptions {
            AccessMode = DocumentAccessMode.ReadOnly,
            OpenSettings = options.OpenXmlMaxCharactersInPart.HasValue
                ? new OpenSettings { MaxCharactersInPart = options.OpenXmlMaxCharactersInPart.Value }
                : null
        };
        if (stream.CanSeek) stream.Position = 0;
        try {
            return ExcelDocument.Load(stream, loadOptions);
        } catch (Exception exception) when (stream.CanSeek && !string.IsNullOrEmpty(options.OpenPassword) && exception is InvalidDataException or IOException) {
            stream.Position = 0;
            return ExcelDocument.LoadEncrypted(stream, options.OpenPassword!, loadOptions);
        }
    }

    private static IEnumerable<ReaderChunk> Extract(
        ExcelDocumentReader reader,
        string sourceName,
        ReaderOptions readerOptions,
        ReaderExcelOptions options,
        IReadOnlyList<string>? legacyWarnings,
        CancellationToken cancellationToken) {
        IReadOnlyList<string> sheets = string.IsNullOrWhiteSpace(options.SheetName)
            ? reader.GetSheetNames()
            : new[] { options.SheetName!.Trim() };
        int tableIndex = 0;
        bool firstChunk = true;
        foreach (string sheet in sheets) {
            foreach (ExcelExtractChunk source in reader.ExtractChunks(
                         sheet,
                         options.A1Range,
                         new ExcelExtractionExtensions.ExcelExtractOptions {
                             HeadersInFirstRow = options.HeadersInFirstRow,
                             ChunkRows = options.ChunkRows,
                             EmitMarkdownTable = true
                         },
                         new ExcelExtractChunkingOptions {
                             MaxChars = readerOptions.MaxChars,
                             MaxTableRows = readerOptions.MaxTableRows
                         },
                         sourceName,
                         cancellationToken)) {
                ReaderTable[] tables = source.Tables.Select(table => new ReaderTable {
                    Title = table.Title,
                    Columns = table.Columns,
                    Rows = table.Rows,
                    TotalRowCount = table.TotalRowCount,
                    Truncated = table.Truncated,
                    ColumnProfiles = ReaderTableProfiler.CreateProfiles(table.Columns, table.Rows),
                    Location = new ReaderLocation {
                        Path = source.Location.Path,
                        Sheet = source.Location.Sheet,
                        A1Range = source.Location.A1Range,
                        SourceBlockIndex = source.Location.BlockIndex,
                        SourceBlockKind = "table",
                        TableIndex = tableIndex++
                    }
                }).ToArray();
                yield return new ReaderChunk {
                    Id = source.Id,
                    Kind = ReaderInputKind.Excel,
                    Location = new ReaderLocation {
                        Path = source.Location.Path,
                        Sheet = source.Location.Sheet,
                        A1Range = source.Location.A1Range,
                        BlockIndex = source.Location.BlockIndex,
                        SourceBlockIndex = source.Location.BlockIndex,
                        SourceBlockKind = "sheet"
                    },
                    Text = source.Text,
                    Markdown = source.Markdown,
                    Tables = tables,
                    Warnings = firstChunk ? Combine(source.Warnings, legacyWarnings) : source.Warnings
                };
                firstChunk = false;
            }
        }
    }

    internal static IReadOnlyList<string>? BuildLegacyWarnings(ExcelDocument document) {
        if (document.SourceFormat != ExcelFileFormat.Xls) return null;

        var warnings = new List<string>();
        AddBounded(warnings, document.LegacyXlsImportDiagnostics.Select(static item => "Legacy XLS import diagnostic: " + item), 8, "Additional legacy XLS import diagnostics were omitted.");
        AddBounded(warnings, document.LegacyXlsUnsupportedFeatures.Select(static item => $"Legacy XLS unsupported feature: {item.Code} ({item.Kind}) - {item.Description}"), 8, "Additional legacy XLS unsupported features were omitted.");
        AddBounded(warnings, document.LegacyXlsPreservedFeatures.Select(static item => $"Legacy XLS preserved feature: {item.Code} ({item.Kind}) - {item.Description}"), 8, "Additional legacy XLS preserved features were omitted.");
        AddBounded(warnings, document.LegacyXlsUnsupportedSheets.Select(static item => $"Legacy XLS unsupported sheet: {item.Name} ({item.Kind}, {item.VisibilityName})"), 8, "Additional legacy XLS unsupported sheets were omitted.");
        AddBounded(warnings, document.LegacyXlsChartSheets.Select(static item => $"Legacy XLS chart sheet: {item.Name} ({item.VisibilityName})"), 8, "Additional legacy XLS chart sheets were omitted.");
        AddBounded(warnings, document.LegacyXlsCompoundFeatures.Select(static item => $"Legacy XLS compound feature: {item.Kind} - entries: {item.Entries.Count}"), 8, "Additional legacy XLS compound features were omitted.");
        return warnings.Count == 0 ? null : warnings;
    }

    private static IReadOnlyList<string>? Combine(IReadOnlyList<string>? first, IReadOnlyList<string>? second) {
        if (first == null || first.Count == 0) return second;
        if (second == null || second.Count == 0) return first;
        return first.Concat(second).ToArray();
    }

    private static void AddBounded(List<string> target, IEnumerable<string> values, int maxItems, string overflowMessage) {
        int count = 0;
        foreach (string value in values) {
            if (count < maxItems && !string.IsNullOrWhiteSpace(value)) target.Add(value);
            count++;
        }
        if (count > maxItems) target.Add(overflowMessage);
    }
}
