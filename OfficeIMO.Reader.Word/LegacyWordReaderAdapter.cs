using OfficeIMO.Word.Legacy;

namespace OfficeIMO.Reader.Word;

internal static class LegacyWordReaderAdapter {
    private const int WordForDosHeaderLength = 97;

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

    internal static bool Probe(Stream stream, string? sourceName, ReaderOptions readerOptions,
        LegacyWordImportOptions? importOptions, CancellationToken cancellationToken) {
        if (!stream.CanSeek) return false;
        long position = stream.Position;
        try {
            LegacyWordImportOptions configured = Prepare(importOptions, sourceName, readerOptions);
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
            LegacyWordImporter.Detect(data, configured, cancellationToken);
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

    internal static bool HasWordForDosHeader(string path, CancellationToken cancellationToken) {
        if (!string.Equals(Path.GetExtension(path), ".doc", StringComparison.OrdinalIgnoreCase)) return false;
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return HasWordForDosHeader(stream, path, cancellationToken);
    }

    internal static bool HasWordForDosHeader(Stream stream, string? sourceName, CancellationToken cancellationToken) {
        if (!string.Equals(Path.GetExtension(sourceName), ".doc", StringComparison.OrdinalIgnoreCase) || !stream.CanSeek) return false;
        long position = stream.Position;
        try {
            var header = new byte[WordForDosHeaderLength];
            int total = 0;
            while (total < header.Length) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = stream.Read(header, total, header.Length - total);
                if (read <= 0) break;
                total += read;
            }
            return total == header.Length && (header[0] == 0x31 || header[0] == 0x32) &&
                   header[1] == 0xBE && header[5] == 0xAB && header[96] == 0;
        } finally {
            stream.Position = position;
        }
    }

    internal static OfficeDocumentReadResult ReadWordForDosDocument(string path, ReaderOptions readerOptions, ReaderWordOptions options,
        LegacyWordImportOptions? importOptions, CancellationToken cancellationToken) {
        LegacyWordImportOptions configured = Prepare(importOptions, path, readerOptions);
        configured.FormatHint = LegacyWordFormat.WordForDos;
        using LegacyWordImportResult imported = LegacyWordImporter.Import(path, configured, cancellationToken);
        return WordReaderAdapter.Project(imported.Document, path, readerOptions, options, cancellationToken, BuildWarnings(imported), OfficeDocumentReaderBuilderWordExtensions.LegacyHandlerId);
    }

    internal static OfficeDocumentReadResult ReadWordForDosDocument(Stream stream, string? sourceName, ReaderOptions readerOptions, ReaderWordOptions options,
        LegacyWordImportOptions? importOptions, CancellationToken cancellationToken) {
        string logicalName = string.IsNullOrWhiteSpace(sourceName) ? "legacy-document.doc" : sourceName!;
        LegacyWordImportOptions configured = Prepare(importOptions, sourceName, readerOptions);
        configured.FormatHint = LegacyWordFormat.WordForDos;
        using LegacyWordImportResult imported = LegacyWordImporter.Import(stream, configured, cancellationToken);
        return WordReaderAdapter.Project(imported.Document, logicalName, readerOptions, options, cancellationToken, BuildWarnings(imported), OfficeDocumentReaderBuilderWordExtensions.LegacyHandlerId);
    }

    private static LegacyWordImportOptions Prepare(LegacyWordImportOptions? source, string? sourceName, ReaderOptions readerOptions) {
        OfficeLegacyImportLimits limits = (source?.Limits ?? new OfficeLegacyImportLimits()).Clone();
        if (readerOptions.MaxInputBytes.HasValue) {
            limits.MaxInputBytes = (int)Math.Min(limits.MaxInputBytes, Math.Min(int.MaxValue, readerOptions.MaxInputBytes.Value));
        }
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
