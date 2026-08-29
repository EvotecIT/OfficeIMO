using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO.Opml;
using OfficeIMO.Reader;
using OfficeIMO.Core.Internal;

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
        OpmlValidationResult validation = document.Validate(null, cancellationToken);
        OpmlConversionResult<OfficeDocumentModel> conversion = document.ToOfficeDocumentModel(
            sourceName, options.ConversionOptions, cancellationToken);
        OpmlDiagnostic[] diagnostics = validation.Diagnostics.Concat(conversion.Diagnostics).ToArray();
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
        cancellationToken.ThrowIfCancellationRequested();
        int sourceIndex = 0, emittedIndex = 0;
        bool warningsAttached = false;
        foreach (OpmlOutline root in document.Outlines) {
            foreach (ReaderChunk chunk in BuildOutline(root, 1, string.Empty)) yield return chunk;
        }
        if (!warningsAttached && warnings != null && warnings.Count > 0) {
            yield return new ReaderChunk {
                Id = "opml-diagnostics",
                Kind = ReaderInputKind.Opml,
                Text = string.Empty,
                Markdown = string.Empty,
                Location = new ReaderLocation { Path = sourceName, BlockIndex = emittedIndex, SourceBlockKind = "diagnostic" },
                Diagnostics = new ReaderChunkDiagnostics { SourceKind = "opml" },
                Warnings = TakeWarnings()
            };
        }

        IReadOnlyList<string>? TakeWarnings() {
            if (warningsAttached || warnings == null) return null;
            warningsAttached = true;
            return warnings;
        }

        IEnumerable<ReaderChunk> BuildOutline(OpmlOutline outline, int level, string parentPath) {
            cancellationToken.ThrowIfCancellationRequested();
            string headingPath = OfficeDocumentHeadingPath.Append(parentPath, outline.Text, " > ");
            int currentSource = sourceIndex++;
            IReadOnlyList<string> parts = DocumentReaderEngine.SplitAdapterProjection(outline.Text, reader.MaxChars);
            if (parts.Count == 0) parts = new[] { string.Empty };
            string targetMarkdown = BuildTargetMarkdown(outline);
            for (int part = 0; part < parts.Count; part++) {
                string markdown = part == 0 ? new string('#', Math.Min(level, 6)) + " " + parts[part] : parts[part];
                if (part == parts.Count - 1 && targetMarkdown.Length > 0) markdown += "\n\n" + targetMarkdown;
                yield return new ReaderChunk {
                    Id = parts.Count == 1 ? "opml-" + currentSource : "opml-" + currentSource + "-part-" + (part + 1),
                    Kind = ReaderInputKind.Opml, Text = parts[part],
                    Markdown = markdown,
                    ContinuesPreviousChunk = part > 0,
                    Location = new ReaderLocation { Path = sourceName, BlockIndex = emittedIndex++, SourceBlockIndex = currentSource,
                        HeadingPath = headingPath, SourceBlockKind = "outline", BlockAnchor = "opml-outline-" + currentSource },
                    Diagnostics = new ReaderChunkDiagnostics { SourceKind = "opml" }, Warnings = TakeWarnings()
                };
            }
            foreach (OpmlOutline child in outline.Children) foreach (ReaderChunk chunk in BuildOutline(child, level + 1, headingPath)) yield return chunk;
        }
    }

    private static string BuildTargetMarkdown(OpmlOutline outline) {
        var targets = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        Add("Feed", outline.XmlUrl);
        Add("Website", outline.HtmlUrl);
        Add("Link", outline.Url);
        return string.Join("\n", targets);

        void Add(string label, string? target) {
            if (string.IsNullOrWhiteSpace(target) || !seen.Add(target!)) return;
            targets.Add("- " + label + ": [" + EscapeMarkdownLabel(target!) + "](" + EscapeMarkdownDestination(target!) + ")");
        }
    }

    private static string EscapeMarkdownLabel(string value) =>
        value.Replace("\\", "\\\\").Replace("[", "\\[").Replace("]", "\\]");

    private static string EscapeMarkdownDestination(string value) {
        var escaped = new System.Text.StringBuilder(value.Length);
        foreach (char character in value) {
            if (char.IsWhiteSpace(character)) {
                foreach (byte utf8Byte in System.Text.Encoding.UTF8.GetBytes(character.ToString())) {
                    escaped.Append('%').Append(utf8Byte.ToString("X2"));
                }
            } else if (character == '\\' || character == '(' || character == ')') {
                escaped.Append('\\').Append(character);
            } else {
                escaped.Append(character);
            }
        }
        return escaped.ToString();
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

}
