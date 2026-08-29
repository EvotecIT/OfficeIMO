using OfficeIMO.Word.Legacy;

namespace OfficeIMO.Reader.Word;

internal static class LegacyWordReaderAdapter {
    internal static LegacyWordImportOptions? Clone(LegacyWordImportOptions? source) {
        if (source == null) return null;
        return new LegacyWordImportOptions {
            Limits = (source.Limits ?? new OfficeLegacyImportLimits()).Clone(),
            FormatHint = source.FormatHint,
            SourceName = source.SourceName,
            RequireStructured = source.RequireStructured
        };
    }

    internal static OfficeDocumentReadResult ReadDocument(string path, ReaderOptions readerOptions, ReaderWordOptions options,
        LegacyWordImportOptions? importOptions, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        using LegacyWordImportResult imported = LegacyWordImporter.Import(path, Prepare(importOptions, path, readerOptions), cancellationToken);
        return WordReaderAdapter.Project(imported.Document, path, readerOptions, options, cancellationToken, BuildWarnings(imported), OfficeDocumentReaderBuilderWordExtensions.LegacyHandlerId);
    }

    internal static OfficeDocumentReadResult ReadDocument(Stream stream, string? sourceName, ReaderOptions readerOptions, ReaderWordOptions options,
        LegacyWordImportOptions? importOptions, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        string logicalName = string.IsNullOrWhiteSpace(sourceName) ? "legacy-document" : sourceName!;
        using LegacyWordImportResult imported = LegacyWordImporter.Import(stream, Prepare(importOptions, sourceName, readerOptions), cancellationToken);
        return WordReaderAdapter.Project(imported.Document, logicalName, readerOptions, options, cancellationToken, BuildWarnings(imported), OfficeDocumentReaderBuilderWordExtensions.LegacyHandlerId);
    }

    private static LegacyWordImportOptions Prepare(LegacyWordImportOptions? source, string? sourceName, ReaderOptions readerOptions) {
        OfficeLegacyImportLimits limits = (source?.Limits ?? new OfficeLegacyImportLimits()).Clone();
        if (readerOptions.MaxInputBytes.HasValue) limits.MaxInputBytes = (int)Math.Min(int.MaxValue, readerOptions.MaxInputBytes.Value);
        return new LegacyWordImportOptions {
            Limits = limits,
            FormatHint = source?.FormatHint,
            SourceName = string.IsNullOrWhiteSpace(source?.SourceName) ? sourceName : source!.SourceName,
            RequireStructured = source?.RequireStructured ?? false
        };
    }

    private static IReadOnlyList<string> BuildWarnings(LegacyWordImportResult imported) {
        var warnings = new List<string> { $"Legacy import quality: {imported.Report.Quality}; profile: {imported.Detection.ProfileId}." };
        warnings.AddRange(imported.Report.Findings.Select(static finding => finding.Code + ": " + finding.Message));
        return warnings;
    }
}
