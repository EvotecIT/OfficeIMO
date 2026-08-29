using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO;
using OfficeIMO.DocBook;
using OfficeIMO.Reader;

namespace OfficeIMO.Reader.DocBook;

internal static partial class DocBookReaderAdapter {
    internal static IEnumerable<ReaderChunk> Read(string path, ReaderOptions? readerOptions = null, ReaderDocBookOptions? docBookOptions = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        ReaderOptions reader = readerOptions ?? new ReaderOptions(); ReaderInputLimits.EnforceFileSize(path, reader.MaxInputBytes);
        ReaderDocBookOptions adapter = ReaderDocBookOptionsCloner.Clone(docBookOptions); ApplyReaderLimit(adapter.ReadOptions, reader.MaxInputBytes);
        cancellationToken.ThrowIfCancellationRequested();
        return Build(DocBookDocument.Load(path, adapter.ReadOptions, cancellationToken), path, reader, adapter, cancellationToken).ToArray();
    }

    internal static IEnumerable<ReaderChunk> Read(Stream stream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderDocBookOptions? docBookOptions = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        ReaderOptions reader = readerOptions ?? new ReaderOptions(); ReaderDocBookOptions adapter = ReaderDocBookOptionsCloner.Clone(docBookOptions);
        ApplyReaderLimit(adapter.ReadOptions, reader.MaxInputBytes);
        return Build(DocBookDocument.Load(stream, adapter.ReadOptions, cancellationToken), string.IsNullOrWhiteSpace(sourceName) ? "document.xml" : sourceName!, reader, adapter, cancellationToken).ToArray();
    }

    internal static IEnumerable<ReaderChunk> Read(DocBookDocument document, string sourceName = "document.xml", ReaderOptions? readerOptions = null, ReaderDocBookOptions? options = null, CancellationToken cancellationToken = default) =>
        Build(document ?? throw new ArgumentNullException(nameof(document)), sourceName, readerOptions ?? new ReaderOptions(), ReaderDocBookOptionsCloner.Clone(options), cancellationToken);

    private static DocBookProjection CreateProjection(DocBookDocument document, string sourceName, ReaderOptions reader, ReaderDocBookOptions options, bool includeChunkWarnings, CancellationToken cancellationToken) {
        DocBookConversionResult<OfficeDocumentModel> conversion = document.ToOfficeDocumentModel(sourceName,
            new DocBookConversionOptions { MaxTableRows = Math.Max(1, reader.MaxTableRows) });
        DocBookDiagnostic[] diagnostics = document.Validate().Diagnostics.Concat(conversion.Diagnostics).ToArray();
        IReadOnlyList<string>? warnings = includeChunkWarnings && options.IncludeDiagnostics
            ? diagnostics.Where(d => d.Severity != DocBookDiagnosticSeverity.Info)
                .Select(d => d.Code + ": " + d.Message).ToArray()
            : null;
        OfficeDocumentModel model = conversion.Value;
        return new DocBookProjection(model, diagnostics,
            BuildChunks(model, sourceName, reader, warnings, cancellationToken).ToArray());
    }

    private static IEnumerable<ReaderChunk> Build(DocBookDocument document, string sourceName, ReaderOptions reader, ReaderDocBookOptions options, CancellationToken cancellationToken) =>
        CreateProjection(document, sourceName, reader, options, includeChunkWarnings: true, cancellationToken).Chunks;

    private static IEnumerable<ReaderChunk> BuildChunks(OfficeDocumentModel model, string sourceName, ReaderOptions reader, IReadOnlyList<string>? warnings, CancellationToken cancellationToken) {
        ReaderTable[] tables = model.Tables.Select(table => MapTable(table, reader.MaxTableRows, sourceName)).ToArray();
        bool tablesAttached = false;
        bool warningsAttached = false;
        int sourceIndex = 0, emittedIndex = 0;
        foreach (OfficeDocumentModelNode root in model.Structure) {
            foreach (ReaderChunk chunk in BuildNode(root)) {
                if (!tablesAttached && tables.Length > 0) {
                    chunk.Tables = tables;
                    tablesAttached = true;
                }
                yield return chunk;
            }
        }
        if (!tablesAttached && tables.Length > 0) {
            yield return new ReaderChunk {
                Id = "docbook-tables",
                Kind = ReaderInputKind.DocBook,
                Text = string.Empty,
                Markdown = string.Empty,
                Tables = tables,
                Location = new ReaderLocation { Path = sourceName, BlockIndex = emittedIndex, SourceBlockKind = "table" },
                Diagnostics = new ReaderChunkDiagnostics { SourceKind = "docbook", TableCount = tables.Length },
                Warnings = TakeWarnings()
            };
        }

        IReadOnlyList<string>? TakeWarnings() {
            if (warningsAttached || warnings == null) return null;
            warningsAttached = true;
            return warnings;
        }

        IEnumerable<ReaderChunk> BuildNode(OfficeDocumentModelNode node) {
            cancellationToken.ThrowIfCancellationRequested();
            int currentSource = sourceIndex++;
            if (!string.IsNullOrWhiteSpace(node.Text) && node.Kind != "metadata" && node.Kind != "author") {
                IReadOnlyList<string> parts = DocumentReaderEngine.SplitAdapterProjection(node.Text, reader.MaxChars);
                for (int part = 0; part < parts.Count; part++) {
                    string markdown;
                    if (node.Kind == "code") {
                        markdown = parts.Count == 1 ? "```\n" + parts[part] + "\n```"
                            : (part == 0 ? "```\n" : string.Empty) + parts[part] +
                              (part == parts.Count - 1 ? "\n```" : string.Empty);
                    } else {
                        markdown = part == 0 && (node.Kind == "section" || node.Kind == "title")
                            ? new string('#', Math.Min(node.Level ?? 1, 6)) + " " + parts[part]
                            : parts[part];
                    }
                    yield return new ReaderChunk {
                        Id = parts.Count == 1 ? "docbook-" + currentSource : "docbook-" + currentSource + "-part-" + (part + 1),
                        Kind = ReaderInputKind.DocBook, Text = parts[part], Markdown = markdown,
                        ContinuesPreviousChunk = part > 0,
                        Location = new ReaderLocation { Path = sourceName, BlockIndex = emittedIndex++, SourceBlockIndex = currentSource,
                            HeadingPath = node.Location.HeadingPath, SourceBlockKind = node.Kind, BlockAnchor = "docbook-node-" + currentSource },
                        Diagnostics = new ReaderChunkDiagnostics { SourceKind = "docbook" }, Warnings = TakeWarnings()
                    };
                }
            }
            bool ownsInlineText = node.Kind == "paragraph" || node.Kind == "code" || node.Kind == "screen" ||
                node.Kind == "title" || node.Kind == "subtitle" || node.Kind == "author" ||
                node.Kind == "link" || node.Kind == "table-cell" || node.Kind == "caption" ||
                (node.Kind.StartsWith("extension:", StringComparison.Ordinal) && !string.IsNullOrWhiteSpace(node.Text)) ||
                (!string.IsNullOrWhiteSpace(node.Text) && node.Children.Count > 0 &&
                 node.Children.All(child => child.Kind == "text") &&
                 string.Equals(string.Concat(node.Children.Select(child => child.Text)), node.Text, StringComparison.Ordinal));
            if (!ownsInlineText) {
                foreach (OfficeDocumentModelNode child in node.Children) {
                    if ((node.Kind == "section" || node.Kind == "table" || node.Kind == "figure") && child.Kind == "title") continue;
                    foreach (ReaderChunk chunk in BuildNode(child)) yield return chunk;
                }
            }
        }
    }

    private sealed class DocBookProjection {
        internal DocBookProjection(OfficeDocumentModel model, IReadOnlyList<DocBookDiagnostic> diagnostics, ReaderChunk[] chunks) {
            Model = model;
            Diagnostics = diagnostics;
            Chunks = chunks;
        }
        internal OfficeDocumentModel Model { get; }
        internal IReadOnlyList<DocBookDiagnostic> Diagnostics { get; }
        internal ReaderChunk[] Chunks { get; }
    }

    private static void ApplyReaderLimit(DocBookReadOptions options, long? maxBytes) {
        if (maxBytes.HasValue) options.MaxInputBytes = Math.Min(options.MaxInputBytes, maxBytes.Value);
    }
    private static ReaderTable MapTable(OfficeDocumentModelTable table, int maxRows, string sourceName) {
        int rowLimit = Math.Max(1, maxRows);
        IReadOnlyList<IReadOnlyList<string>> rows = table.Rows.Take(rowLimit).ToArray();
        return new ReaderTable {
            Title = table.Title,
            Kind = table.Kind,
            Summary = table.Summary,
            PayloadHash = table.PayloadHash,
            Location = MapLocation(table.Location, sourceName, "table"),
            Columns = table.Columns,
            ColumnProfiles = ReaderTableProfiler.CreateProfiles(table.Columns, rows),
            Rows = rows,
            TotalRowCount = Math.Max(table.TotalRowCount, table.Rows.Count),
            Truncated = table.Truncated || rows.Count < table.Rows.Count
        };
    }
    private static ReaderLocation MapLocation(OfficeDocumentModelLocation? location, string sourceName, string? defaultSourceBlockKind = null) => new ReaderLocation {
        Path = location?.Path ?? sourceName,
        BlockIndex = location?.BlockIndex,
        SourceBlockIndex = location?.SourceBlockIndex,
        StartLine = location?.StartLine,
        EndLine = location?.EndLine,
        NormalizedStartLine = location?.NormalizedStartLine,
        NormalizedEndLine = location?.NormalizedEndLine,
        HeadingPath = location?.HeadingPath,
        HeadingSlug = location?.HeadingSlug,
        SourceBlockKind = location?.SourceBlockKind ?? defaultSourceBlockKind,
        BlockAnchor = location?.BlockAnchor,
        Sheet = location?.Sheet,
        A1Range = location?.A1Range,
        Slide = location?.Slide,
        Page = location?.Page,
        TableIndex = location?.TableIndex
    };
}
