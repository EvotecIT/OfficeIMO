using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    /// <summary>
    /// Reads a supported document file and returns discovered visual payloads in source order.
    /// </summary>
    /// <param name="path">Source file path.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<ReaderVisual> ReadVisuals(string path, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        return ExtractVisuals(Read(path, options, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Reads a supported document stream and returns discovered visual payloads in source order.
    /// </summary>
    /// <param name="stream">Source stream. This method does not close the stream.</param>
    /// <param name="sourceName">Optional source name used for kind detection and citations.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<ReaderVisual> ReadVisuals(Stream stream, string? sourceName = null, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        return ExtractVisuals(Read(stream, sourceName, options, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Reads a supported document from bytes and returns discovered visual payloads in source order.
    /// </summary>
    /// <param name="bytes">Source bytes.</param>
    /// <param name="sourceName">Optional source name used for kind detection and citations.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<ReaderVisual> ReadVisuals(byte[] bytes, string? sourceName = null, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using var stream = new MemoryStream(bytes, writable: false);
        return ReadVisuals(stream, sourceName, options, cancellationToken);
    }

    /// <summary>
    /// Reads a supported document file and returns discovered visuals with deterministic payload and JSON sidecars.
    /// </summary>
    /// <param name="path">Source file path.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="indentedJson">When true, writes indented visual JSON payloads.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<ReaderVisualExportBundle> ReadVisualExports(string path, ReaderOptions? options = null, bool indentedJson = false, CancellationToken cancellationToken = default) {
        return ExportVisuals(ReadVisuals(path, options, cancellationToken), indentedJson, cancellationToken);
    }

    /// <summary>
    /// Reads a supported document stream and returns discovered visuals with deterministic payload and JSON sidecars.
    /// </summary>
    /// <param name="stream">Source stream. This method does not close the stream.</param>
    /// <param name="sourceName">Optional source name used for kind detection and citations.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="indentedJson">When true, writes indented visual JSON payloads.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<ReaderVisualExportBundle> ReadVisualExports(Stream stream, string? sourceName = null, ReaderOptions? options = null, bool indentedJson = false, CancellationToken cancellationToken = default) {
        return ExportVisuals(ReadVisuals(stream, sourceName, options, cancellationToken), indentedJson, cancellationToken);
    }

    /// <summary>
    /// Reads a supported document from bytes and returns discovered visuals with deterministic payload and JSON sidecars.
    /// </summary>
    /// <param name="bytes">Source bytes.</param>
    /// <param name="sourceName">Optional source name used for kind detection and citations.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="indentedJson">When true, writes indented visual JSON payloads.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<ReaderVisualExportBundle> ReadVisualExports(byte[] bytes, string? sourceName = null, ReaderOptions? options = null, bool indentedJson = false, CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using var stream = new MemoryStream(bytes, writable: false);
        return ReadVisualExports(stream, sourceName, options, indentedJson, cancellationToken);
    }

    /// <summary>
    /// Extracts visual payload metadata already attached to reader chunks.
    /// </summary>
    /// <param name="chunks">Reader chunks to inspect.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<ReaderVisual> ExtractVisuals(IEnumerable<ReaderChunk> chunks, CancellationToken cancellationToken = default) {
        if (chunks == null) throw new ArgumentNullException(nameof(chunks));

        List<ReaderVisual>? visuals = null;
        foreach (ReaderChunk? chunk in chunks) {
            cancellationToken.ThrowIfCancellationRequested();
            if (chunk?.Visuals == null || chunk.Visuals.Count == 0) {
                continue;
            }

            visuals ??= new List<ReaderVisual>();
            for (int i = 0; i < chunk.Visuals.Count; i++) {
                visuals.Add(WithChunkLocationFallback(chunk.Visuals[i], chunk.Location));
            }
        }

        return visuals == null || visuals.Count == 0 ? Array.Empty<ReaderVisual>() : visuals.ToArray();
    }

    /// <summary>
    /// Builds deterministic payload and JSON sidecars for discovered reader visuals.
    /// </summary>
    /// <param name="visuals">Reader visuals to export.</param>
    /// <param name="indentedJson">When true, writes indented visual JSON payloads.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<ReaderVisualExportBundle> ExportVisuals(IEnumerable<ReaderVisual> visuals, bool indentedJson = false, CancellationToken cancellationToken = default) {
        if (visuals == null) throw new ArgumentNullException(nameof(visuals));

        List<ReaderVisualExportBundle>? exports = null;
        int index = 0;
        foreach (ReaderVisual? visual in visuals) {
            cancellationToken.ThrowIfCancellationRequested();
            if (visual == null) {
                continue;
            }

            exports ??= new List<ReaderVisualExportBundle>();
            exports.Add(BuildVisualExport(visual, index, indentedJson));
            index++;
        }

        return exports == null || exports.Count == 0 ? Array.Empty<ReaderVisualExportBundle>() : exports.ToArray();
    }

    private static ReaderVisual WithChunkLocationFallback(ReaderVisual visual, ReaderLocation chunkLocation) {
        return new ReaderVisual {
            Kind = visual.Kind,
            Language = visual.Language,
            Content = visual.Content,
            PayloadHash = visual.PayloadHash,
            Location = MergeVisualLocation(visual.Location, chunkLocation),
            SourceName = visual.SourceName,
            MimeType = visual.MimeType,
            Width = visual.Width,
            Height = visual.Height,
            X = visual.X,
            Y = visual.Y,
            PlacedWidth = visual.PlacedWidth,
            PlacedHeight = visual.PlacedHeight,
            PlacementCount = visual.PlacementCount,
            HasGeometry = visual.HasGeometry,
            IsAxisAligned = visual.IsAxisAligned
        };
    }

    private static ReaderLocation MergeVisualLocation(ReaderLocation? visualLocation, ReaderLocation chunkLocation) {
        return new ReaderLocation {
            Path = string.IsNullOrWhiteSpace(visualLocation?.Path) ? chunkLocation.Path : visualLocation!.Path,
            BlockIndex = visualLocation?.BlockIndex ?? chunkLocation.BlockIndex,
            SourceBlockIndex = visualLocation?.SourceBlockIndex ?? chunkLocation.SourceBlockIndex,
            StartLine = visualLocation?.StartLine ?? chunkLocation.StartLine,
            EndLine = visualLocation?.EndLine ?? chunkLocation.EndLine,
            NormalizedStartLine = visualLocation?.NormalizedStartLine ?? chunkLocation.NormalizedStartLine,
            NormalizedEndLine = visualLocation?.NormalizedEndLine ?? chunkLocation.NormalizedEndLine,
            HeadingPath = visualLocation?.HeadingPath ?? chunkLocation.HeadingPath,
            HeadingSlug = visualLocation?.HeadingSlug ?? chunkLocation.HeadingSlug,
            SourceBlockKind = visualLocation?.SourceBlockKind ?? chunkLocation.SourceBlockKind,
            BlockAnchor = visualLocation?.BlockAnchor ?? chunkLocation.BlockAnchor,
            Sheet = visualLocation?.Sheet ?? chunkLocation.Sheet,
            A1Range = visualLocation?.A1Range ?? chunkLocation.A1Range,
            Slide = visualLocation?.Slide ?? chunkLocation.Slide,
            Page = visualLocation?.Page ?? chunkLocation.Page,
            TableIndex = visualLocation?.TableIndex ?? chunkLocation.TableIndex
        };
    }

    private static ReaderVisualExportBundle BuildVisualExport(ReaderVisual visual, int index, bool indentedJson) {
        string id = BuildVisualExportId(visual, index);
        return new ReaderVisualExportBundle {
            Id = id,
            FileNamePrefix = OfficeDocumentAssetNaming.BuildFileName(id, null),
            PayloadExtension = ReaderVisualExport.GetPayloadExtension(visual),
            Visual = visual,
            Payload = visual.Content ?? string.Empty,
            Json = visual.ToJson(indentedJson)
        };
    }

    private static string BuildVisualExportId(ReaderVisual visual, int index) {
        ReaderLocation? location = visual.Location;
        string container = location == null ? "document" : BuildVisualLocationStem(location);
        return container + "-visual-" + index.ToString("D4", CultureInfo.InvariantCulture);
    }

    private static string BuildVisualLocationStem(ReaderLocation location) {
        if (!string.IsNullOrWhiteSpace(location.Path)) {
            string path = Path.GetFileNameWithoutExtension(location.Path) ?? location.Path!;
            if (!string.IsNullOrWhiteSpace(path)) {
                return path + BuildVisualLocationSuffix(location);
            }
        }

        return "document" + BuildVisualLocationSuffix(location);
    }

    private static string BuildVisualLocationSuffix(ReaderLocation location) {
        if (!string.IsNullOrWhiteSpace(location.Sheet)) {
            return "-sheet-" + location.Sheet;
        }

        if (location.Page.HasValue) {
            return "-page-" + location.Page.Value.ToString("D4", CultureInfo.InvariantCulture);
        }

        if (location.Slide.HasValue) {
            return "-slide-" + location.Slide.Value.ToString("D4", CultureInfo.InvariantCulture);
        }

        return string.Empty;
    }
}
