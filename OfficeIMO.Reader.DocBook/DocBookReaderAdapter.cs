using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO;
using OfficeIMO.DocBook;
using OfficeIMO.Reader;

namespace OfficeIMO.Reader.DocBook;

internal static class DocBookReaderAdapter {
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

    private static IEnumerable<ReaderChunk> Build(DocBookDocument document, string sourceName, ReaderOptions reader, ReaderDocBookOptions options, CancellationToken cancellationToken) {
        IReadOnlyList<string>? warnings = options.IncludeDiagnostics
            ? document.Validate().Diagnostics.Where(d => d.Severity != DocBookDiagnosticSeverity.Info).Select(d => d.Code + ": " + d.Message).ToArray() : null;
        OfficeDocumentModel model = document.ToOfficeDocumentModel(sourceName).Value;
        int sourceIndex = 0, emittedIndex = 0;
        foreach (OfficeDocumentModelNode root in model.Structure) foreach (ReaderChunk chunk in BuildNode(root)) yield return chunk;

        IEnumerable<ReaderChunk> BuildNode(OfficeDocumentModelNode node) {
            cancellationToken.ThrowIfCancellationRequested();
            int currentSource = sourceIndex++;
            if (!string.IsNullOrWhiteSpace(node.Text) && node.Kind != "metadata" && node.Kind != "author") {
                IReadOnlyList<string> parts = Split(node.Text, reader.MaxChars);
                for (int part = 0; part < parts.Count; part++) {
                    string markdown = node.Kind == "section" || node.Kind == "title"
                        ? new string('#', Math.Min(node.Level ?? 1, 6)) + " " + parts[part]
                        : node.Kind == "code" ? "```\n" + parts[part] + "\n```" : parts[part];
                    yield return new ReaderChunk {
                        Id = parts.Count == 1 ? "docbook-" + currentSource : "docbook-" + currentSource + "-part-" + (part + 1),
                        Kind = ReaderInputKind.DocBook, Text = parts[part], Markdown = markdown,
                        Location = new ReaderLocation { Path = sourceName, BlockIndex = emittedIndex++, SourceBlockIndex = currentSource,
                            HeadingPath = node.Location.HeadingPath, SourceBlockKind = node.Kind, BlockAnchor = "docbook-node-" + currentSource },
                        Diagnostics = new ReaderChunkDiagnostics { SourceKind = "docbook" }, Warnings = warnings
                    };
                }
            }
            bool ownsInlineText = node.Kind == "paragraph" || node.Kind == "code" || node.Kind == "screen" ||
                node.Kind == "title" || node.Kind == "subtitle" || node.Kind == "author" ||
                node.Kind == "link" || node.Kind == "table-cell" || node.Kind == "caption" ||
                node.Kind.StartsWith("extension:", StringComparison.Ordinal) && !string.IsNullOrWhiteSpace(node.Text);
            if (!ownsInlineText) {
                foreach (OfficeDocumentModelNode child in node.Children) {
                    if ((node.Kind == "section" || node.Kind == "table" || node.Kind == "figure") && child.Kind == "title") continue;
                    foreach (ReaderChunk chunk in BuildNode(child)) yield return chunk;
                }
            }
        }
    }

    private static void ApplyReaderLimit(DocBookReadOptions options, long? maxBytes) {
        if (maxBytes.HasValue) options.MaxInputBytes = Math.Min(options.MaxInputBytes, maxBytes.Value);
    }
    private static IReadOnlyList<string> Split(string value, int maxChars) {
        if (maxChars <= 0 || value.Length <= maxChars) return new[] { value };
        var parts = new List<string>();
        for (int offset = 0; offset < value.Length; offset += maxChars) parts.Add(value.Substring(offset, Math.Min(maxChars, value.Length - offset)));
        return parts;
    }
}
