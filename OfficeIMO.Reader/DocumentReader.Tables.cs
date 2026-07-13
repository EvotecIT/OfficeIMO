using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    /// <summary>
    /// Reads a supported document file and returns discovered tables in source order.
    /// </summary>
    /// <param name="path">Source file path.</param>
    /// <param name="options">Extraction options. <see cref="ReaderOptions.MaxTableRows"/> is honored.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<ReaderTable> ReadTables(string path, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        return ExtractTables(Read(path, options, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Reads a supported document stream and returns discovered tables in source order.
    /// </summary>
    /// <param name="stream">Source stream. This method does not close the stream.</param>
    /// <param name="sourceName">Optional source name used for kind detection and citations.</param>
    /// <param name="options">Extraction options. <see cref="ReaderOptions.MaxTableRows"/> is honored.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<ReaderTable> ReadTables(Stream stream, string? sourceName = null, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        return ExtractTables(Read(stream, sourceName, options, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Reads a supported document from bytes and returns discovered tables in source order.
    /// </summary>
    /// <param name="bytes">Source bytes.</param>
    /// <param name="sourceName">Optional source name used for kind detection and citations.</param>
    /// <param name="options">Extraction options. <see cref="ReaderOptions.MaxTableRows"/> is honored.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<ReaderTable> ReadTables(byte[] bytes, string? sourceName = null, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using var stream = new MemoryStream(bytes, writable: false);
        return ReadTables(stream, sourceName, options, cancellationToken);
    }

    /// <summary>
    /// Reads a supported document file and returns discovered tables with deterministic CSV, Markdown, and JSON payloads.
    /// </summary>
    /// <param name="path">Source file path.</param>
    /// <param name="options">Extraction options. <see cref="ReaderOptions.MaxTableRows"/> is honored.</param>
    /// <param name="indentedJson">When true, writes indented table JSON payloads.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<ReaderTableExportBundle> ReadTableExports(string path, ReaderOptions? options = null, bool indentedJson = false, CancellationToken cancellationToken = default) {
        return ExportTables(ReadTables(path, options, cancellationToken), indentedJson, cancellationToken);
    }

    /// <summary>
    /// Reads a supported document stream and returns discovered tables with deterministic CSV, Markdown, and JSON payloads.
    /// </summary>
    /// <param name="stream">Source stream. This method does not close the stream.</param>
    /// <param name="sourceName">Optional source name used for kind detection and citations.</param>
    /// <param name="options">Extraction options. <see cref="ReaderOptions.MaxTableRows"/> is honored.</param>
    /// <param name="indentedJson">When true, writes indented table JSON payloads.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<ReaderTableExportBundle> ReadTableExports(Stream stream, string? sourceName = null, ReaderOptions? options = null, bool indentedJson = false, CancellationToken cancellationToken = default) {
        return ExportTables(ReadTables(stream, sourceName, options, cancellationToken), indentedJson, cancellationToken);
    }

    /// <summary>
    /// Reads a supported document from bytes and returns discovered tables with deterministic CSV, Markdown, and JSON payloads.
    /// </summary>
    /// <param name="bytes">Source bytes.</param>
    /// <param name="sourceName">Optional source name used for kind detection and citations.</param>
    /// <param name="options">Extraction options. <see cref="ReaderOptions.MaxTableRows"/> is honored.</param>
    /// <param name="indentedJson">When true, writes indented table JSON payloads.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<ReaderTableExportBundle> ReadTableExports(byte[] bytes, string? sourceName = null, ReaderOptions? options = null, bool indentedJson = false, CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using var stream = new MemoryStream(bytes, writable: false);
        return ReadTableExports(stream, sourceName, options, indentedJson, cancellationToken);
    }

    /// <summary>
    /// Extracts table metadata already attached to reader chunks.
    /// </summary>
    /// <param name="chunks">Reader chunks to inspect.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<ReaderTable> ExtractTables(IEnumerable<ReaderChunk> chunks, CancellationToken cancellationToken = default) {
        if (chunks == null) throw new ArgumentNullException(nameof(chunks));

        List<ReaderTable>? tables = null;
        foreach (ReaderChunk? chunk in chunks) {
            cancellationToken.ThrowIfCancellationRequested();
            if (chunk?.Tables == null || chunk.Tables.Count == 0) {
                continue;
            }

            tables ??= new List<ReaderTable>();
            for (int i = 0; i < chunk.Tables.Count; i++) {
                tables.Add(WithChunkLocationFallback(chunk.Tables[i], chunk.Location, tables.Count));
            }
        }

        return tables == null || tables.Count == 0 ? Array.Empty<ReaderTable>() : tables.ToArray();
    }

    /// <summary>
    /// Builds deterministic CSV, Markdown, and JSON payloads for discovered reader tables.
    /// </summary>
    /// <param name="tables">Reader tables to export.</param>
    /// <param name="indentedJson">When true, writes indented table JSON payloads.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<ReaderTableExportBundle> ExportTables(IEnumerable<ReaderTable> tables, bool indentedJson = false, CancellationToken cancellationToken = default) {
        if (tables == null) throw new ArgumentNullException(nameof(tables));

        List<ReaderTableExportBundle>? exports = null;
        int index = 0;
        foreach (ReaderTable? table in tables) {
            cancellationToken.ThrowIfCancellationRequested();
            if (table == null) {
                continue;
            }

            exports ??= new List<ReaderTableExportBundle>();
            exports.Add(BuildTableExport(table, index, indentedJson));
            index++;
        }

        return exports == null || exports.Count == 0 ? Array.Empty<ReaderTableExportBundle>() : exports.ToArray();
    }

    private static ReaderTable WithChunkLocationFallback(ReaderTable table, ReaderLocation chunkLocation, int fallbackTableIndex) {
        if (table.Location != null && !NeedsLocationFallback(table.Location)) {
            return table;
        }

        return new ReaderTable {
            Title = table.Title,
            Kind = table.Kind,
            CallId = table.CallId,
            Summary = table.Summary,
            PayloadHash = table.PayloadHash,
            Location = MergeLocation(table.Location, chunkLocation, fallbackTableIndex),
            Columns = table.Columns,
            ColumnProfiles = table.ColumnProfiles,
            Diagnostics = table.Diagnostics,
            Rows = table.Rows,
            TotalRowCount = table.TotalRowCount,
            Truncated = table.Truncated
        };
    }

    private static bool NeedsLocationFallback(ReaderLocation location) {
        return string.IsNullOrWhiteSpace(location.Path)
            || (!location.Page.HasValue && !location.Slide.HasValue && string.IsNullOrWhiteSpace(location.Sheet));
    }

    private static ReaderLocation MergeLocation(ReaderLocation? tableLocation, ReaderLocation chunkLocation, int fallbackTableIndex) {
        return new ReaderLocation {
            Path = string.IsNullOrWhiteSpace(tableLocation?.Path) ? chunkLocation.Path : tableLocation!.Path,
            BlockIndex = tableLocation?.BlockIndex ?? chunkLocation.BlockIndex,
            SourceBlockIndex = tableLocation?.SourceBlockIndex ?? chunkLocation.SourceBlockIndex,
            StartLine = tableLocation?.StartLine ?? chunkLocation.StartLine,
            EndLine = tableLocation?.EndLine ?? chunkLocation.EndLine,
            NormalizedStartLine = tableLocation?.NormalizedStartLine ?? chunkLocation.NormalizedStartLine,
            NormalizedEndLine = tableLocation?.NormalizedEndLine ?? chunkLocation.NormalizedEndLine,
            HeadingPath = tableLocation?.HeadingPath ?? chunkLocation.HeadingPath,
            HeadingSlug = tableLocation?.HeadingSlug ?? chunkLocation.HeadingSlug,
            SourceBlockKind = tableLocation?.SourceBlockKind ?? chunkLocation.SourceBlockKind,
            BlockAnchor = tableLocation?.BlockAnchor ?? chunkLocation.BlockAnchor,
            Sheet = tableLocation?.Sheet ?? chunkLocation.Sheet,
            A1Range = tableLocation?.A1Range ?? chunkLocation.A1Range,
            Slide = tableLocation?.Slide ?? chunkLocation.Slide,
            Page = tableLocation?.Page ?? chunkLocation.Page,
            TableIndex = tableLocation?.TableIndex ?? fallbackTableIndex
        };
    }

    private static ReaderTableExportBundle BuildTableExport(ReaderTable table, int index, bool indentedJson) {
        string id = BuildTableExportId(table, index);
        return new ReaderTableExportBundle {
            Id = id,
            FileNamePrefix = OfficeDocumentAssetNaming.BuildFileName(id, null),
            Table = table,
            Csv = table.ToCsv(),
            Markdown = table.ToMarkdownTable(),
            Json = table.ToJson(indentedJson)
        };
    }

    private static string BuildTableExportId(ReaderTable table, int index) {
        ReaderLocation? location = table.Location;
        string container = location == null ? "document" : BuildTableLocationStem(location);
        int tableIndex = location?.TableIndex ?? index;
        return container + BuildTableSelectionSuffix(location) + "-table-" + tableIndex.ToString("D4", CultureInfo.InvariantCulture);
    }

    private static string BuildTableLocationStem(ReaderLocation location) {
        if (!string.IsNullOrWhiteSpace(location.Path)) {
            string path = Path.GetFileNameWithoutExtension(location.Path) ?? location.Path!;
            if (!string.IsNullOrWhiteSpace(path)) {
                return path + BuildTableLocationSuffix(location);
            }
        }

        return "document" + BuildTableLocationSuffix(location);
    }

    private static string BuildTableLocationSuffix(ReaderLocation location) {
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

    private static string BuildTableSelectionSuffix(ReaderLocation? location) {
        if (location?.Page == null || location.SourceBlockIndex == null || location.SourceBlockIndex.Value <= 0) {
            return string.Empty;
        }

        return "-selection-" + location.SourceBlockIndex.Value.ToString("D4", CultureInfo.InvariantCulture);
    }
}
