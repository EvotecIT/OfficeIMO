using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.Reader.FormatInternals;

namespace OfficeIMO.Reader.PowerPoint;

internal static class PowerPointReaderAdapter {
    internal const long DefaultModernMaxInputBytes = 512L * 1024L * 1024L;

    internal static ReaderPowerPointOptions Clone(ReaderPowerPointOptions? source) => new ReaderPowerPointOptions {
        IncludeNotes = source?.IncludeNotes ?? true,
        IncludeTables = source?.IncludeTables ?? true,
        IncludeHiddenShapes = source?.IncludeHiddenShapes ?? true
    };

    internal static OfficeDocumentReadResult ReadDocument(string path, ReaderOptions readerOptions, ReaderPowerPointOptions options, CancellationToken cancellationToken) {
        using PowerPointDocument presentation = Load(path, readerOptions, cancellationToken);
        return Project(presentation, path, readerOptions, options, cancellationToken);
    }

    internal static OfficeDocumentReadResult ReadDocument(Stream stream, string? sourceName, ReaderOptions readerOptions, ReaderPowerPointOptions options, CancellationToken cancellationToken) {
        using PowerPointDocument presentation = Load(stream, sourceName, readerOptions, cancellationToken);
        return Project(presentation, string.IsNullOrWhiteSpace(sourceName) ? "presentation.pptx" : sourceName!, readerOptions, options, cancellationToken);
    }

    internal static bool ProbeEncryptedOpenXml(Stream stream, string? sourceName, ReaderOptions options, CancellationToken cancellationToken) {
        if (string.IsNullOrEmpty(options.OpenPassword) || !stream.CanSeek) return false;
        long position = stream.Position;
        try {
            cancellationToken.ThrowIfCancellationRequested();
            using PowerPointDocument presentation = Load(stream, sourceName, options, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            return presentation.OpenXmlDocument.GetAllParts().Any(static part =>
                string.Equals(part.Uri.OriginalString, "/ppt/presentation.xml",
                    StringComparison.OrdinalIgnoreCase));
        } catch (OperationCanceledException) {
            throw;
        } catch (ReaderResourceLimitException) {
            throw;
        } catch (OfficePackageSecurityException) {
            throw;
        } catch {
            return false;
        } finally {
            stream.Position = position;
        }
    }

    private static PowerPointDocument Load(string path, ReaderOptions options, CancellationToken cancellationToken) =>
        OpenXmlReaderAdmission.Load(options, () => PowerPointDocument.Load(path, CreateLoadOptions(options, path), cancellationToken));

    private static PowerPointDocument Load(Stream stream, string? sourceName, ReaderOptions options, CancellationToken cancellationToken) {
        if (stream.CanSeek) stream.Position = 0;
        return OpenXmlReaderAdmission.Load(options, () => PowerPointDocument.Load(stream, CreateLoadOptions(options, sourceName), cancellationToken));
    }

    internal static PowerPointLoadOptions CreateLoadOptions(ReaderOptions options, string? sourceName) {
        long defaultMaxInputBytes = IsLegacySourceName(sourceName)
            ? LegacyPptImportOptions.DefaultMaxInputBytes
            : DefaultModernMaxInputBytes;
        long maxInputBytes = options.MaxInputBytes ?? defaultMaxInputBytes;
        long legacyMaxInputBytes = options.MaxInputBytes ?? LegacyPptImportOptions.DefaultMaxInputBytes;
        return new PowerPointLoadOptions {
            AccessMode = DocumentAccessMode.ReadOnly,
            AllowLegacyProjection = options.Projection == ReaderProjection.Full,
            CaptureSourcePackageSecurity = options.Inspection?.IncludePackageContent == true,
            OpenSettings = options.OpenXmlMaxCharactersInPart.HasValue
                ? new OfficeOpenXmlLoadSettings { MaxCharactersInPart = options.OpenXmlMaxCharactersInPart.Value }
                : null,
            PackageSecurity = OpenXmlReaderAdmission.Policy(options, maxInputBytes, requirePolicy: true),
            MaxInputBytes = maxInputBytes,
            MaxLegacyInputBytes = legacyMaxInputBytes > int.MaxValue ? int.MaxValue : checked((int)legacyMaxInputBytes),
            Password = options.OpenPassword,
            ReportUnsupportedContent = true
        };
    }

    private static bool IsLegacySourceName(string? sourceName) {
        string extension = Path.GetExtension(sourceName ?? string.Empty);
        return PowerPointFormatCatalog.All.Any(format =>
            format.Generation == OfficeFormatGeneration.Legacy &&
            string.Equals(format.Extension, extension, StringComparison.OrdinalIgnoreCase));
    }

internal static OfficeDocumentReadResult Project(PowerPointDocument presentation, string sourceName, ReaderOptions readerOptions, ReaderPowerPointOptions options, CancellationToken cancellationToken) {
        IReadOnlyList<string>? legacyWarnings = BuildLegacyWarnings(presentation);
        IReadOnlyList<OfficeDocumentAsset> assets = readerOptions.Projection == ReaderProjection.TextAndMetadata ? Array.Empty<OfficeDocumentAsset>() : OpenXmlImageAssetCollector.CollectPowerPoint(
            presentation.OpenXmlDocument, sourceName, readerOptions, options.IncludeNotes,
            options.IncludeHiddenShapes, cancellationToken);
        var chunks = presentation.ExtractMarkdownChunks(
                new PowerPointExtractionExtensions.PowerPointExtractOptions {
                    IncludeNotes = options.IncludeNotes,
                    IncludeTables = options.IncludeTables,
                    IncludeHiddenShapes = options.IncludeHiddenShapes,
                    IncludeMarkdown = readerOptions.Projection == ReaderProjection.Full
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
        result = PowerPointRichMapping.Apply(presentation, readerOptions, options, result, cancellationToken);
        return PowerPointInspectionProjection.Apply(presentation, readerOptions, result, cancellationToken);
    }

    private static IReadOnlyList<string>? Combine(IReadOnlyList<string>? first, IReadOnlyList<string>? second) {
        if (first == null || first.Count == 0) return second;
        if (second == null || second.Count == 0) return first;
        return first.Concat(second).ToArray();
    }

    private static IReadOnlyList<string>? BuildLegacyWarnings(PowerPointDocument presentation) {
        if (presentation.SourceFormat is not PowerPointFileFormat.Ppt and not PowerPointFileFormat.Pot and not PowerPointFileFormat.Pps) return null;
        string[] warnings = presentation.LegacyPptImportDiagnostics
            .Select(static diagnostic => "Legacy PPT import diagnostic: " + diagnostic)
            .Take(16)
            .ToArray();
        return warnings.Length == 0 ? null : warnings;
    }
}
