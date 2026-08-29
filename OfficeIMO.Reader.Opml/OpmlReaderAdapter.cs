using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO.Opml;
using OfficeIMO.Reader;

namespace OfficeIMO.Reader.Opml;

internal static partial class OpmlReaderAdapter {
    internal static IEnumerable<ReaderChunk> Read(string path, ReaderOptions? readerOptions = null, ReaderOpmlOptions? opmlOptions = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        ReaderOptions reader = readerOptions ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(path, reader.MaxInputBytes);
        ReaderOpmlOptions adapter = ReaderOpmlOptionsCloner.Clone(opmlOptions);
        ApplyReaderLimit(adapter.ReadOptions, reader.MaxInputBytes);
        cancellationToken.ThrowIfCancellationRequested();
        return Build(OpmlDocument.Load(path, adapter.ReadOptions, cancellationToken), path, reader, adapter, cancellationToken).ToArray();
    }

    internal static IEnumerable<ReaderChunk> Read(Stream stream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderOpmlOptions? opmlOptions = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        ReaderOptions reader = readerOptions ?? new ReaderOptions();
        ReaderOpmlOptions adapter = ReaderOpmlOptionsCloner.Clone(opmlOptions);
        ApplyReaderLimit(adapter.ReadOptions, reader.MaxInputBytes);
        return Build(OpmlDocument.Load(stream, adapter.ReadOptions, cancellationToken), string.IsNullOrWhiteSpace(sourceName) ? "document.opml" : sourceName!, reader, adapter, cancellationToken).ToArray();
    }

    internal static IEnumerable<ReaderChunk> Read(OpmlDocument document, string sourceName = "document.opml", ReaderOptions? readerOptions = null, ReaderOpmlOptions? opmlOptions = null, CancellationToken cancellationToken = default) =>
        Build(document ?? throw new ArgumentNullException(nameof(document)), sourceName, readerOptions ?? new ReaderOptions(), ReaderOpmlOptionsCloner.Clone(opmlOptions), cancellationToken);

    private static OpmlProjection CreateProjection(OpmlDocument document, string sourceName, ReaderOptions reader, ReaderOpmlOptions options, bool includeChunkWarnings, CancellationToken cancellationToken) {
        OpmlValidationResult validation = document.Validate();
        OpmlConversionResult<OfficeDocumentModel> conversion = document.ToOfficeDocumentModel(sourceName);
        OpmlDiagnostic[] diagnostics = options.IncludeDiagnostics
            ? validation.Diagnostics.Concat(conversion.Diagnostics).ToArray()
            : Array.Empty<OpmlDiagnostic>();
        IReadOnlyList<string>? warnings = includeChunkWarnings && options.IncludeDiagnostics
            ? diagnostics.Where(d => d.Severity != OpmlDiagnosticSeverity.Info)
                .Select(d => d.Code + ": " + d.Message).ToArray()
            : null;
        return new OpmlProjection(conversion.Value, diagnostics,
            BuildChunks(document, sourceName, reader, warnings, cancellationToken).ToArray());
    }

    private static IEnumerable<ReaderChunk> Build(OpmlDocument document, string sourceName, ReaderOptions reader, ReaderOpmlOptions options, CancellationToken cancellationToken) =>
        CreateProjection(document, sourceName, reader, options, includeChunkWarnings: true, cancellationToken).Chunks;

    private static IEnumerable<ReaderChunk> BuildChunks(OpmlDocument document, string sourceName, ReaderOptions reader, IReadOnlyList<string>? warnings, CancellationToken cancellationToken) {
        int sourceIndex = 0, emittedIndex = 0;
        foreach (OpmlOutline root in document.Outlines) {
            foreach (ReaderChunk chunk in BuildOutline(root, 1, string.Empty)) yield return chunk;
        }

        IEnumerable<ReaderChunk> BuildOutline(OpmlOutline outline, int level, string parentPath) {
            cancellationToken.ThrowIfCancellationRequested();
            string headingPath = string.IsNullOrEmpty(parentPath) ? outline.Text : parentPath + " > " + outline.Text;
            int currentSource = sourceIndex++;
            IReadOnlyList<string> parts = Split(outline.Text, reader.MaxChars);
            if (parts.Count == 0) parts = new[] { string.Empty };
            for (int part = 0; part < parts.Count; part++) {
                yield return new ReaderChunk {
                    Id = parts.Count == 1 ? "opml-" + currentSource : "opml-" + currentSource + "-part-" + (part + 1),
                    Kind = ReaderInputKind.Opml, Text = parts[part], Markdown = new string('#', Math.Min(level, 6)) + " " + parts[part],
                    Location = new ReaderLocation { Path = sourceName, BlockIndex = emittedIndex++, SourceBlockIndex = currentSource,
                        HeadingPath = headingPath, SourceBlockKind = "outline", BlockAnchor = "opml-outline-" + currentSource },
                    Diagnostics = new ReaderChunkDiagnostics { SourceKind = "opml" }, Warnings = warnings
                };
            }
            foreach (OpmlOutline child in outline.Children) foreach (ReaderChunk chunk in BuildOutline(child, level + 1, headingPath)) yield return chunk;
        }
    }

    private sealed class OpmlProjection {
        internal OpmlProjection(OfficeDocumentModel model, IReadOnlyList<OpmlDiagnostic> diagnostics, ReaderChunk[] chunks) {
            Model = model;
            Diagnostics = diagnostics;
            Chunks = chunks;
        }
        internal OfficeDocumentModel Model { get; }
        internal IReadOnlyList<OpmlDiagnostic> Diagnostics { get; }
        internal ReaderChunk[] Chunks { get; }
    }

    private static void ApplyReaderLimit(OpmlReadOptions options, long? maxBytes) {
        if (maxBytes.HasValue) options.MaxInputBytes = Math.Min(options.MaxInputBytes, maxBytes.Value);
    }

    private static IReadOnlyList<string> Split(string value, int maxChars) {
        if (value.Length == 0 || maxChars <= 0 || value.Length <= maxChars) return new[] { value };
        var parts = new List<string>();
        for (int offset = 0; offset < value.Length; offset += maxChars) parts.Add(value.Substring(offset, Math.Min(maxChars, value.Length - offset)));
        return parts;
    }
}
