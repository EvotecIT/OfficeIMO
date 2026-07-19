using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    /// <summary>
    /// Reads a supported document file and returns the shared OfficeIMO read result envelope.
    /// </summary>
    /// <param name="path">Source file path.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static OfficeDocumentReadResult ReadDocument(string path, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (Directory.Exists(path)) {
            throw new IOException($"'{path}' is a directory. Use {nameof(ReadFolder)}(...) to ingest directories.");
        }
        if (!File.Exists(path)) throw new FileNotFoundException($"File '{path}' doesn't exist.", path);

        ReaderOptions opt = NormalizeOptions(options);
        EnforceFileSize(path, ResolveInitialMaxInputBytes(path, opt));
        if (!TryResolvePathHandler(path, opt, cancellationToken,
                out ReaderHandlerDescriptor handler, out ReaderDetectionResult detection)) {
            throw CreateUnsupportedInputException(path, detection);
        }
        if (handler.ReadDocumentPath != null) {
            OfficeDocumentReadResult result = ValidateDocumentResult(handler.ReadDocumentPath(path, opt, cancellationToken), handler.Id);
            SourceInfo source = BuildSourceInfoFromPath(path, ShouldComputeSourceHash(handler, opt), cancellationToken);
            return ApplyDetectionDiagnostics(FinalizeHandlerDocumentResult(result, source, opt.ComputeHashes), detection);
        }

        ReaderChunk[] chunks = Read(path, opt, cancellationToken).ToArray();
        return BuildChunkDocumentResult(
            chunks, path, handler.Kind, BuildPathDocumentSource(path, chunks), detection: detection);
    }

    /// <summary>
    /// Reads a supported document stream and returns the shared OfficeIMO read result envelope.
    /// </summary>
    /// <param name="stream">Source stream. This method does not close the stream.</param>
    /// <param name="sourceName">Optional source name used for kind detection and citations.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static OfficeDocumentReadResult ReadDocument(Stream stream, string? sourceName = null, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        ReaderOptions opt = NormalizeOptions(options);
        string logicalSourceName = NormalizeLogicalSourceName(sourceName, "memory");
        Stream readStream = ReaderInputLimits.EnsureSeekableReadStream(
            stream,
            ResolveStreamMaxInputBytes(logicalSourceName, opt,
                stream.CanSeek),
            cancellationToken,
            out bool ownsReadStream);
        try {
            bool hasCustomStreamHandler = TryResolveStreamHandler(
                readStream,
                logicalSourceName,
                opt,
                cancellationToken,
                out ReaderHandlerDescriptor customStreamHandler,
                out ReaderDetectionResult detection);
            if (!hasCustomStreamHandler) {
                throw CreateUnsupportedInputException(logicalSourceName, detection);
            }
            if (customStreamHandler.ReadDocumentStream != null) {
                OfficeDocumentReadResult result = ValidateDocumentResult(
                    customStreamHandler.ReadDocumentStream(readStream, logicalSourceName, opt, cancellationToken),
                    customStreamHandler.Id);
                SourceInfo source = BuildSourceInfoFromStream(readStream, logicalSourceName,
                    ShouldComputeSourceHash(customStreamHandler, opt), cancellationToken);
                return ApplyDetectionDiagnostics(FinalizeHandlerDocumentResult(result, source, opt.ComputeHashes), detection);
            }

            long position = readStream.Position;
            ReaderChunk[] chunks = Read(readStream, logicalSourceName, opt, cancellationToken).ToArray();
            if (readStream.CanSeek) readStream.Position = position;
            return BuildChunkDocumentResult(
                chunks,
                logicalSourceName,
                customStreamHandler.Kind,
                BuildStreamDocumentSource(readStream, logicalSourceName, chunks),
                detection: detection);
        } finally {
            if (ownsReadStream) {
                readStream.Dispose();
            }
        }
    }

    private static IReadOnlyList<ReaderChunk> GetDocumentResultChunks(OfficeDocumentReadResult? result, string handlerId) {
        return ValidateDocumentResult(result, handlerId).Chunks ?? Array.Empty<ReaderChunk>();
    }

    private static OfficeDocumentReadResult ValidateDocumentResult(OfficeDocumentReadResult? result, string handlerId) {
        if (result == null) {
            throw new InvalidOperationException($"Reader handler '{handlerId}' returned a null document result.");
        }

        return result;
    }

    private static OfficeDocumentReadResult FinalizeHandlerDocumentResult(OfficeDocumentReadResult result, SourceInfo source, bool computeHashes) {
        result.Source ??= new OfficeDocumentSource();
        result.Source.Path ??= source.Path;
        result.Source.SourceId ??= source.SourceId;
        result.Source.SourceHash ??= source.SourceHash;
        result.Source.LastWriteUtc ??= source.LastWriteUtc;
        result.Source.LengthBytes ??= source.LengthBytes;

        IReadOnlyList<ReaderChunk> chunks = result.Chunks ?? Array.Empty<ReaderChunk>();
        for (int index = 0; index < chunks.Count; index++) EnrichChunk(chunks[index], source, computeHashes);
        return result;
    }

    private static bool ShouldComputeSourceHash(ReaderHandlerDescriptor handler, ReaderOptions options) =>
        options.ComputeHashes && handler.SourceHashBehavior == ReaderSourceHashBehavior.InheritReaderOptions;

    /// <summary>
    /// Reads a supported document from bytes and returns the shared OfficeIMO read result envelope.
    /// </summary>
    /// <param name="bytes">Source bytes.</param>
    /// <param name="sourceName">Optional source name used for kind detection and citations.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static OfficeDocumentReadResult ReadDocument(byte[] bytes, string? sourceName = null, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using var ms = new MemoryStream(bytes, writable: false);
        return ReadDocument(ms, sourceName, options, cancellationToken);
    }

    /// <summary>
    /// Reads a supported document file and returns the shared OfficeIMO read result JSON envelope.
    /// </summary>
    public static string ReadDocumentJson(string path, ReaderOptions? options = null, bool indented = false, CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(ReadDocument(path, options, cancellationToken), indented);
    }

    /// <summary>
    /// Reads a supported document stream and returns the shared OfficeIMO read result JSON envelope.
    /// </summary>
    public static string ReadDocumentJson(Stream stream, string? sourceName = null, ReaderOptions? options = null, bool indented = false, CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(ReadDocument(stream, sourceName, options, cancellationToken), indented);
    }

    /// <summary>
    /// Reads a supported document from bytes and returns the shared OfficeIMO read result JSON envelope.
    /// </summary>
    public static string ReadDocumentJson(byte[] bytes, string? sourceName = null, ReaderOptions? options = null, bool indented = false, CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(ReadDocument(bytes, sourceName, options, cancellationToken), indented);
    }

    private static OfficeDocumentReadResult BuildChunkDocumentResult(
        IReadOnlyList<ReaderChunk> chunks,
        string sourceName,
        ReaderInputKind fallbackKind,
        OfficeDocumentSource source,
        IReadOnlyList<OfficeDocumentAsset>? assets = null,
        ReaderDetectionResult? detection = null,
        IReadOnlyList<OfficeDocumentOcrCandidate>? ocrCandidates = null) {
        ReaderInputKind kind = chunks.Count > 0 ? chunks[0].Kind : fallbackKind;
        ReaderTable[] tables = ExtractTables(chunks).ToArray();
        ReaderVisual[] visuals = ExtractVisuals(chunks).ToArray();
        OfficeDocumentAsset[] assetArray = assets == null || assets.Count == 0 ? Array.Empty<OfficeDocumentAsset>() : assets.ToArray();
        OfficeDocumentBlock[] blocks = BuildChunkDocumentBlocks(chunks).ToArray();
        OfficeDocumentOcrCandidate[] candidateArray = ocrCandidates == null
            ? BuildChunkDocumentOcrCandidates(blocks, assetArray).ToArray()
            : ocrCandidates.ToArray();
        IReadOnlyList<OfficeDocumentPage> pages = BuildChunkDocumentPages(chunks, blocks, tables, assetArray, candidateArray);

        return new OfficeDocumentReadResult {
            Kind = kind,
            Source = source,
            CapabilitiesUsed = BuildChunkDocumentCapabilities(kind),
            Markdown = BuildChunkDocumentMarkdown(chunks),
            Chunks = chunks,
            Metadata = BuildChunkDocumentMetadata(kind, chunks, blocks, tables, visuals, pages, assetArray),
            Pages = pages,
            Blocks = blocks,
            Tables = tables,
            Visuals = visuals,
            Diagnostics = BuildChunkDocumentDiagnostics(chunks, candidateArray, detection),
            Assets = assetArray,
            Links = Array.Empty<OfficeDocumentLink>(),
            Forms = Array.Empty<OfficeDocumentFormField>(),
            OcrCandidates = candidateArray
        };
    }

    private static IReadOnlyList<string> BuildChunkDocumentCapabilities(ReaderInputKind kind) {
        if (kind == ReaderInputKind.Unknown) {
            return new[] { "officeimo.reader" };
        }

        return new[] {
            "officeimo.reader",
            "officeimo.reader." + kind.ToString().ToLowerInvariant()
        };
    }

    private static OfficeDocumentSource BuildPathDocumentSource(string path, IReadOnlyList<ReaderChunk> chunks) {
        ReaderChunk? first = chunks.Count > 0 ? chunks[0] : null;
        DateTime? lastWriteUtc = first?.SourceLastWriteUtc;
        long? lengthBytes = first?.SourceLengthBytes;

        if (!lastWriteUtc.HasValue || !lengthBytes.HasValue) {
            try {
                var info = new FileInfo(path);
                if (info.Exists) {
                    lastWriteUtc ??= info.LastWriteTimeUtc;
                    lengthBytes ??= info.Length;
                }
            } catch {
                // Best-effort source metadata.
            }
        }

        return new OfficeDocumentSource {
            Path = first?.Location.Path ?? path,
            SourceId = first?.SourceId ?? BuildSourceId(NormalizePathForId(path)),
            SourceHash = first?.SourceHash,
            LastWriteUtc = lastWriteUtc,
            LengthBytes = lengthBytes
        };
    }

    private static OfficeDocumentSource BuildStreamDocumentSource(Stream stream, string sourceName, IReadOnlyList<ReaderChunk> chunks) {
        ReaderChunk? first = chunks.Count > 0 ? chunks[0] : null;
        long? lengthBytes = first?.SourceLengthBytes;

        if (!lengthBytes.HasValue) {
            try {
                if (stream.CanSeek) {
                    lengthBytes = stream.Length;
                }
            } catch {
                // Best-effort source metadata.
            }
        }

        return new OfficeDocumentSource {
            Path = first?.Location.Path ?? sourceName,
            SourceId = first?.SourceId ?? BuildSourceId(sourceName),
            SourceHash = first?.SourceHash,
            LastWriteUtc = first?.SourceLastWriteUtc,
            LengthBytes = lengthBytes
        };
    }

    private static string? BuildChunkDocumentMarkdown(IReadOnlyList<ReaderChunk> chunks) {
        return JoinChunkMarkdown(
            chunks,
            static (chunk, _) => string.IsNullOrWhiteSpace(chunk.Markdown) ? chunk.Text : chunk.Markdown);
    }

    private static string? JoinChunkMarkdown(
        IReadOnlyList<ReaderChunk> chunks,
        Func<ReaderChunk, int, string?> valueSelector) {
        StringBuilder? markdown = null;
        for (int index = 0; index < chunks.Count; index++) {
            ReaderChunk chunk = chunks[index];
            string? value = valueSelector(chunk, index);
            if (string.IsNullOrEmpty(value) ||
                (string.IsNullOrWhiteSpace(value) && !chunk.ContinuesPreviousChunk)) continue;

            if (markdown == null) {
                markdown = new StringBuilder(value!.Length);
            } else if (!chunk.ContinuesPreviousChunk) {
                markdown.AppendLine().AppendLine();
            }
            markdown.Append(value);
        }
        return markdown?.ToString();
    }

    private static IEnumerable<OfficeDocumentBlock> BuildChunkDocumentBlocks(IReadOnlyList<ReaderChunk> chunks) {
        for (int i = 0; i < chunks.Count; i++) {
            ReaderChunk chunk = chunks[i];
            string blockId = !string.IsNullOrWhiteSpace(chunk.Id)
                ? chunk.Id
                : "chunk-" + i.ToString("D4", CultureInfo.InvariantCulture);
            yield return new OfficeDocumentBlock {
                Id = blockId,
                Kind = string.IsNullOrWhiteSpace(chunk.Location.SourceBlockKind) ? "chunk" : chunk.Location.SourceBlockKind!,
                Text = chunk.Text ?? string.Empty,
                Location = chunk.Location
            };
        }
    }

    private static IReadOnlyList<OfficeDocumentPage> BuildChunkDocumentPages(
        IReadOnlyList<ReaderChunk> chunks,
        IReadOnlyList<OfficeDocumentBlock> blocks,
        IReadOnlyList<ReaderTable> tables,
        IReadOnlyList<OfficeDocumentAsset> assets,
        IReadOnlyList<OfficeDocumentOcrCandidate> ocrCandidates) {
        if (chunks.Count == 0 && tables.Count == 0 && assets.Count == 0 && ocrCandidates.Count == 0) {
            return Array.Empty<OfficeDocumentPage>();
        }

        var pages = new List<OfficeDocumentPage>();
        var pageNumbers = chunks
            .Where(static chunk => chunk.Location.Page.HasValue)
            .Select(static chunk => chunk.Location.Page!.Value)
            .Concat(tables.Where(static table => table.Location?.Page != null).Select(static table => table.Location!.Page!.Value))
            .Concat(assets.Where(static asset => asset.Location.Page.HasValue).Select(static asset => asset.Location.Page!.Value))
            .Concat(ocrCandidates.Where(static candidate => candidate.Location.Page.HasValue).Select(static candidate => candidate.Location.Page!.Value))
            .Distinct()
            .OrderBy(static page => page);
        foreach (int pageNumber in pageNumbers) {
            ReaderLocation location = chunks.FirstOrDefault(chunk => chunk.Location.Page == pageNumber)?.Location
                ?? tables.FirstOrDefault(table => table.Location?.Page == pageNumber)?.Location
                ?? assets.FirstOrDefault(asset => asset.Location.Page == pageNumber)?.Location
                ?? ocrCandidates.First(candidate => candidate.Location.Page == pageNumber).Location;
            pages.Add(BuildChunkPage(
                number: pageNumber,
                name: null,
                location: BuildContainerLocation(location, "page"),
                blocks: blocks.Where(block => block.Location.Page == pageNumber).ToArray(),
                tables: tables.Where(table => table.Location?.Page == pageNumber).ToArray(),
                assets: assets.Where(asset => asset.Location.Page == pageNumber).ToArray(),
                ocrCandidates: ocrCandidates.Where(candidate => candidate.Location.Page == pageNumber).ToArray()));
        }

        var slideNumbers = chunks
            .Where(static chunk => chunk.Location.Slide.HasValue)
            .Select(static chunk => chunk.Location.Slide!.Value)
            .Concat(tables.Where(static table => table.Location?.Slide != null).Select(static table => table.Location!.Slide!.Value))
            .Concat(assets.Where(static asset => asset.Location.Slide.HasValue).Select(static asset => asset.Location.Slide!.Value))
            .Concat(ocrCandidates.Where(static candidate => candidate.Location.Slide.HasValue).Select(static candidate => candidate.Location.Slide!.Value))
            .Distinct()
            .OrderBy(static slide => slide);
        foreach (int slideNumber in slideNumbers) {
            ReaderLocation location = chunks.FirstOrDefault(chunk => chunk.Location.Slide == slideNumber)?.Location
                ?? tables.FirstOrDefault(table => table.Location?.Slide == slideNumber)?.Location
                ?? assets.FirstOrDefault(asset => asset.Location.Slide == slideNumber)?.Location
                ?? ocrCandidates.First(candidate => candidate.Location.Slide == slideNumber).Location;
            pages.Add(BuildChunkPage(
                number: slideNumber,
                name: null,
                location: BuildContainerLocation(location, "slide"),
                blocks: blocks.Where(block => block.Location.Slide == slideNumber).ToArray(),
                tables: tables.Where(table => table.Location?.Slide == slideNumber).ToArray(),
                assets: assets.Where(asset => asset.Location.Slide == slideNumber).ToArray(),
                ocrCandidates: ocrCandidates.Where(candidate => candidate.Location.Slide == slideNumber).ToArray()));
        }

        var sheetNames = chunks
            .Where(static chunk => !string.IsNullOrWhiteSpace(chunk.Location.Sheet))
            .Select(static chunk => chunk.Location.Sheet!)
            .Concat(tables.Where(static table => !string.IsNullOrWhiteSpace(table.Location?.Sheet)).Select(static table => table.Location!.Sheet!))
            .Concat(assets.Where(static asset => !string.IsNullOrWhiteSpace(asset.Location.Sheet)).Select(static asset => asset.Location.Sheet!))
            .Concat(ocrCandidates.Where(static candidate => !string.IsNullOrWhiteSpace(candidate.Location.Sheet)).Select(static candidate => candidate.Location.Sheet!))
            .Distinct(StringComparer.Ordinal);
        foreach (string sheetName in sheetNames) {
            ReaderLocation location = chunks.FirstOrDefault(chunk => string.Equals(chunk.Location.Sheet, sheetName, StringComparison.Ordinal))?.Location
                ?? tables.FirstOrDefault(table => string.Equals(table.Location?.Sheet, sheetName, StringComparison.Ordinal))?.Location
                ?? assets.FirstOrDefault(asset => string.Equals(asset.Location.Sheet, sheetName, StringComparison.Ordinal))?.Location
                ?? ocrCandidates.First(candidate => string.Equals(candidate.Location.Sheet, sheetName, StringComparison.Ordinal)).Location;
            pages.Add(BuildChunkPage(
                number: null,
                name: sheetName,
                location: BuildContainerLocation(location, "sheet"),
                blocks: blocks.Where(block => string.Equals(block.Location.Sheet, sheetName, StringComparison.Ordinal)).ToArray(),
                tables: tables.Where(table => string.Equals(table.Location?.Sheet, sheetName, StringComparison.Ordinal)).ToArray(),
                assets: assets.Where(asset => string.Equals(asset.Location.Sheet, sheetName, StringComparison.Ordinal)).ToArray(),
                ocrCandidates: ocrCandidates.Where(candidate => string.Equals(candidate.Location.Sheet, sheetName, StringComparison.Ordinal)).ToArray()));
        }

        return pages.Count == 0 ? Array.Empty<OfficeDocumentPage>() : pages;
    }

    private static OfficeDocumentPage BuildChunkPage(int? number, string? name, ReaderLocation location, IReadOnlyList<OfficeDocumentBlock> blocks, IReadOnlyList<ReaderTable> tables, IReadOnlyList<OfficeDocumentAsset> assets, IReadOnlyList<OfficeDocumentOcrCandidate> ocrCandidates) {
        return new OfficeDocumentPage {
            Number = number,
            Name = name,
            Location = location,
            Blocks = blocks,
            Tables = tables,
            Assets = assets,
            Links = Array.Empty<OfficeDocumentLink>(),
            Forms = Array.Empty<OfficeDocumentFormField>(),
            OcrCandidates = ocrCandidates
        };
    }

    private static ReaderLocation BuildContainerLocation(ReaderLocation source, string containerKind) {
        return new ReaderLocation {
            Path = source.Path,
            Sheet = source.Sheet,
            Slide = source.Slide,
            Page = source.Page,
            SourceBlockKind = containerKind,
            BlockAnchor = containerKind + "-" + (source.Page?.ToString(CultureInfo.InvariantCulture) ?? source.Slide?.ToString(CultureInfo.InvariantCulture) ?? source.Sheet ?? "0")
        };
    }

    private static IEnumerable<OfficeDocumentOcrCandidate> BuildChunkDocumentOcrCandidates(IReadOnlyList<OfficeDocumentBlock> blocks, IReadOnlyList<OfficeDocumentAsset> assets) {
        for (int i = 0; i < assets.Count; i++) {
            OfficeDocumentAsset asset = assets[i];
            if (!string.Equals(asset.Kind, "image", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            int textBlockCount = CountSubstantiveTextBlocksForAssetContainer(blocks, asset.Location);
            string reason = textBlockCount == 0
                ? "Image asset has no substantive native text in the same document container."
                : "Image asset may contain text that native extraction cannot inspect.";

            yield return new OfficeDocumentOcrCandidate {
                Id = asset.Id + "-ocr",
                Kind = "image",
                Reason = reason,
                Confidence = textBlockCount == 0 ? 0.75D : 0.35D,
                AssetId = asset.Id,
                ImageCount = 1,
                TextBlockCount = textBlockCount,
                Location = asset.Location
            };
        }
    }

    private static int CountSubstantiveTextBlocksForAssetContainer(IReadOnlyList<OfficeDocumentBlock> blocks, ReaderLocation location) {
        IEnumerable<OfficeDocumentBlock> candidates = blocks;
        if (location.Page.HasValue) {
            candidates = candidates.Where(block => block.Location.Page == location.Page);
        } else if (location.Slide.HasValue) {
            candidates = candidates.Where(block => block.Location.Slide == location.Slide);
        } else if (!string.IsNullOrWhiteSpace(location.Sheet)) {
            candidates = candidates.Where(block => string.Equals(block.Location.Sheet, location.Sheet, StringComparison.Ordinal));
        }

        return candidates.Count(static block => HasSubstantiveNativeText(block.Text));
    }

    private static bool HasSubstantiveNativeText(string? text) {
        if (string.IsNullOrWhiteSpace(text)) {
            return false;
        }

        string value = text!.Trim();
        if (value.StartsWith("![", StringComparison.Ordinal)) {
            return false;
        }

        if (string.Equals(value, "Picture", StringComparison.OrdinalIgnoreCase) ||
            value.StartsWith("Picture ", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        return true;
    }

    private static IReadOnlyList<OfficeDocumentDiagnostic> BuildChunkDocumentDiagnostics(
        IReadOnlyList<ReaderChunk> chunks,
        IReadOnlyList<OfficeDocumentOcrCandidate> ocrCandidates,
        ReaderDetectionResult? detection) {
        var diagnostics = new List<OfficeDocumentDiagnostic>();
        AddDetectionDiagnostic(diagnostics, detection);
        for (int i = 0; i < chunks.Count; i++) {
            ReaderChunk chunk = chunks[i];
            if (chunk.Warnings == null) {
                continue;
            }

            for (int warningIndex = 0; warningIndex < chunk.Warnings.Count; warningIndex++) {
                diagnostics.Add(BuildWarningDiagnostic(chunk.Warnings[warningIndex], chunk.Location));
            }
        }

        for (int i = 0; i < ocrCandidates.Count; i++) {
            OfficeDocumentOcrCandidate candidate = ocrCandidates[i];
            diagnostics.Add(new OfficeDocumentDiagnostic {
                Severity = OfficeDocumentDiagnosticSeverity.Warning,
                Category = OfficeDocumentDiagnosticCategory.Ocr,
                Code = "ocr-needed",
                Message = candidate.Reason ?? "OCR may be needed before text extraction is complete.",
                Source = "officeimo.reader",
                IsRecoverable = true,
                Location = candidate.Location
            });
        }

        return diagnostics.Count == 0 ? Array.Empty<OfficeDocumentDiagnostic>() : diagnostics;
    }

    private static OfficeDocumentDiagnostic BuildWarningDiagnostic(string warning, ReaderLocation location) {
        string message = warning ?? string.Empty;
        string lower = message.ToLowerInvariant();
        string code = "reader-warning";
        OfficeDocumentDiagnosticCategory category = OfficeDocumentDiagnosticCategory.Adapter;

        if (lower.Contains("maxinputbytes") || lower.Contains("maxtotalbytes")) {
            code = "input-limit-exceeded";
            category = OfficeDocumentDiagnosticCategory.Limit;
        } else if (lower.Contains("maxreturnedchunks")) {
            code = "output-limit-reached";
            category = OfficeDocumentDiagnosticCategory.Limit;
        } else if (lower.Contains("parse error") || lower.Contains("malformed")) {
            code = "parse-failed";
            category = OfficeDocumentDiagnosticCategory.Parsing;
        } else if (lower.Contains("unsupported")) {
            code = "unsupported-content";
            category = OfficeDocumentDiagnosticCategory.Content;
        } else if (lower.Contains("truncat") || lower.Contains("split due to maxchars")) {
            code = "content-truncated";
            category = OfficeDocumentDiagnosticCategory.Content;
        } else if (lower.Contains("read error") || lower.Contains("i/o") || lower.Contains("could not be read")) {
            code = "read-failed";
            category = OfficeDocumentDiagnosticCategory.Input;
        }

        return new OfficeDocumentDiagnostic {
            Severity = OfficeDocumentDiagnosticSeverity.Warning,
            Category = category,
            Code = code,
            Message = message,
            Source = "officeimo.reader",
            IsRecoverable = true,
            Location = location
        };
    }

    private static void AddDetectionDiagnostic(
        List<OfficeDocumentDiagnostic> diagnostics,
        ReaderDetectionResult? detection) {
        if (detection == null || !detection.ContentInspected) return;

        if (detection.IsMismatch) {
            diagnostics.Add(new OfficeDocumentDiagnostic {
                Severity = OfficeDocumentDiagnosticSeverity.Warning,
                Category = OfficeDocumentDiagnosticCategory.Detection,
                Code = "input-kind-mismatch",
                Message = $"Content was detected as {detection.ContentKind} but the source extension indicates {detection.ExtensionKind}.",
                Source = "officeimo.reader.detection",
                IsRecoverable = true,
                Location = new ReaderLocation { Path = detection.SourceName },
                Attributes = BuildDetectionAttributes(detection)
            });
        } else if (detection.ExtensionKind == ReaderInputKind.Unknown && detection.ContentKind != ReaderInputKind.Unknown) {
            diagnostics.Add(new OfficeDocumentDiagnostic {
                Severity = OfficeDocumentDiagnosticSeverity.Information,
                Category = OfficeDocumentDiagnosticCategory.Detection,
                Code = "input-kind-detected",
                Message = $"Input kind {detection.ContentKind} was selected from content evidence.",
                Source = "officeimo.reader.detection",
                IsRecoverable = true,
                Location = new ReaderLocation { Path = detection.SourceName },
                Attributes = BuildDetectionAttributes(detection)
            });
        }
    }

    private static IReadOnlyDictionary<string, string> BuildDetectionAttributes(ReaderDetectionResult detection) {
        return new Dictionary<string, string>(StringComparer.Ordinal) {
            ["extensionKind"] = detection.ExtensionKind.ToString(),
            ["contentKind"] = detection.ContentKind.ToString(),
            ["effectiveKind"] = detection.Kind.ToString(),
            ["contentConfidence"] = detection.ContentConfidence.ToString(),
            ["mediaType"] = detection.MediaType ?? string.Empty,
            ["evidence"] = string.Join(",", detection.Evidence)
        };
    }

    private static OfficeDocumentReadResult ApplyDetectionDiagnostics(
        OfficeDocumentReadResult result,
        ReaderDetectionResult detection) {
        var diagnostics = new List<OfficeDocumentDiagnostic>(result.Diagnostics ?? Array.Empty<OfficeDocumentDiagnostic>());
        AddDetectionDiagnostic(diagnostics, detection);
        result.Diagnostics = diagnostics.Count == 0 ? Array.Empty<OfficeDocumentDiagnostic>() : diagnostics.ToArray();
        return result;
    }

    private static string NormalizeLogicalSourceName(string? sourceName, string fallback) {
        if (!string.IsNullOrWhiteSpace(sourceName)) {
            return sourceName!.Trim();
        }

        return fallback;
    }
}
