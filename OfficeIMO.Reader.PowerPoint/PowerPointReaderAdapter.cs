using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.Reader.FormatInternals;

namespace OfficeIMO.Reader.PowerPoint;

internal static class PowerPointReaderAdapter {
    internal static ReaderPowerPointOptions Clone(ReaderPowerPointOptions? source) => new ReaderPowerPointOptions {
        IncludeNotes = source?.IncludeNotes ?? true,
        IncludeTables = source?.IncludeTables ?? true,
        IncludeHiddenShapes = source?.IncludeHiddenShapes ?? true
    };

    internal static OfficeDocumentReadResult ReadDocument(string path, ReaderOptions readerOptions, ReaderPowerPointOptions options, CancellationToken cancellationToken) {
        using PowerPointPresentation presentation = Load(path, readerOptions, cancellationToken);
        return Project(presentation, path, readerOptions, options, cancellationToken);
    }

    internal static OfficeDocumentReadResult ReadDocument(Stream stream, string? sourceName, ReaderOptions readerOptions, ReaderPowerPointOptions options, CancellationToken cancellationToken) {
        using PowerPointPresentation presentation = Load(stream, readerOptions, cancellationToken);
        return Project(presentation, string.IsNullOrWhiteSpace(sourceName) ? "presentation.pptx" : sourceName!, readerOptions, options, cancellationToken);
    }

    internal static bool ProbeEncryptedOpenXml(Stream stream, ReaderOptions options, CancellationToken cancellationToken) {
        if (string.IsNullOrEmpty(options.OpenPassword) || !stream.CanSeek) return false;
        long position = stream.Position;
        try {
            cancellationToken.ThrowIfCancellationRequested();
            using PowerPointPresentation presentation = Load(stream, options, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            return presentation.OpenXmlDocument.GetAllParts().Any(static part =>
                string.Equals(part.Uri.OriginalString, "/ppt/presentation.xml",
                    StringComparison.OrdinalIgnoreCase));
        } catch (OperationCanceledException) {
            throw;
        } catch {
            return false;
        } finally {
            stream.Position = position;
        }
    }

    private static PowerPointPresentation Load(string path, ReaderOptions options, CancellationToken cancellationToken) {
        PowerPointLoadOptions loadOptions = CreateLoadOptions(options);
        try {
            return PowerPointPresentation.Load(path, loadOptions, cancellationToken);
        } catch (Exception exception) when (ShouldRetryEncrypted(exception, options)) {
            try {
                return PowerPointPresentation.LoadEncrypted(path, options.OpenPassword!, loadOptions, cancellationToken);
            } catch (CryptographicException) {
                throw;
            }
        }
    }

    private static PowerPointPresentation Load(Stream stream, ReaderOptions options, CancellationToken cancellationToken) {
        if (stream.CanSeek) stream.Position = 0;
        PowerPointLoadOptions loadOptions = CreateLoadOptions(options);
        try {
            return PowerPointPresentation.Load(stream, loadOptions, cancellationToken);
        } catch (Exception exception) when (stream.CanSeek && ShouldRetryEncrypted(exception, options)) {
            stream.Position = 0;
            try {
                return PowerPointPresentation.LoadEncrypted(stream, options.OpenPassword!, loadOptions, cancellationToken);
            } catch (CryptographicException) {
                throw;
            }
        }
    }

    private static PowerPointLoadOptions CreateLoadOptions(ReaderOptions options) {
        long maxInputBytes = options.MaxInputBytes ?? LegacyPptImportOptions.DefaultMaxInputBytes;
        return new PowerPointLoadOptions {
            AccessMode = DocumentAccessMode.ReadOnly,
            LegacyPptImportOptions = new LegacyPptImportOptions {
                MaxInputBytes = maxInputBytes > int.MaxValue ? int.MaxValue : checked((int)maxInputBytes),
                Password = options.OpenPassword,
                ReportUnsupportedContent = true
            }
        };
    }

    private static bool ShouldRetryEncrypted(Exception exception, ReaderOptions options) =>
        !string.IsNullOrEmpty(options.OpenPassword)
        && exception is InvalidDataException or IOException;

    private static OfficeDocumentReadResult Project(PowerPointPresentation presentation, string sourceName, ReaderOptions readerOptions, ReaderPowerPointOptions options, CancellationToken cancellationToken) {
        IReadOnlyList<string>? legacyWarnings = BuildLegacyWarnings(presentation);
        IReadOnlyList<OfficeDocumentAsset> assets = OpenXmlImageAssetCollector.CollectPowerPoint(
            presentation.OpenXmlDocument, sourceName, readerOptions, options.IncludeNotes,
            options.IncludeHiddenShapes, cancellationToken);
        var chunks = presentation.ExtractMarkdownChunks(
                new PowerPointExtractionExtensions.PowerPointExtractOptions {
                    IncludeNotes = options.IncludeNotes,
                    IncludeTables = options.IncludeTables,
                    IncludeHiddenShapes = options.IncludeHiddenShapes
                },
                new PowerPointExtractChunkingOptions { MaxChars = readerOptions.MaxChars },
                sourceName,
                cancellationToken)
            .Select((chunk, index) => new ReaderChunk {
                Id = chunk.Id,
                Kind = ReaderInputKind.PowerPoint,
                Location = new ReaderLocation {
                    Path = sourceName,
                    Slide = chunk.Location.Slide,
                    BlockIndex = index,
                    SourceBlockIndex = chunk.Location.BlockIndex,
                    SourceBlockKind = "slide"
                },
                Text = chunk.Text,
                Markdown = chunk.Markdown,
                Warnings = Combine(chunk.Warnings, index == 0 ? legacyWarnings : null)
            })
            .ToArray();

        OfficeDocumentReadResult result = DocumentReaderEngine.CreateDocumentResult(
            chunks,
            ReaderInputKind.PowerPoint,
            source: null,
            capabilities: new[] { OfficeDocumentReaderBuilderPowerPointExtensions.HandlerId },
            assets: assets);
        return PowerPointRichMapping.Apply(presentation, readerOptions, options, result, cancellationToken);
    }

    private static IReadOnlyList<string>? Combine(IReadOnlyList<string>? first, IReadOnlyList<string>? second) {
        if (first == null || first.Count == 0) return second;
        if (second == null || second.Count == 0) return first;
        return first.Concat(second).ToArray();
    }

    private static IReadOnlyList<string>? BuildLegacyWarnings(PowerPointPresentation presentation) {
        if (presentation.SourceFormat is not PowerPointFileFormat.Ppt and not PowerPointFileFormat.Pot and not PowerPointFileFormat.Pps) return null;
        string[] warnings = presentation.LegacyPptImportDiagnostics
            .Select(static diagnostic => "Legacy PPT import diagnostic: " + diagnostic)
            .Take(16)
            .ToArray();
        return warnings.Length == 0 ? null : warnings;
    }
}
