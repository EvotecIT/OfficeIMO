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
            int maxChars = Math.Max(1, reader.MaxChars);
            string headingPrefix = new string('#', Math.Min(level, 6)) + " ";
            IReadOnlyList<HeadingProjectionPart> parts = SplitHeadingProjection(outline.Text, headingPrefix, maxChars);
            string targetMarkdown = BuildTargetMarkdown(outline);
            string finalTextMarkdown = parts[parts.Count - 1].Markdown;
            bool appendTargets = targetMarkdown.Length > 0 &&
                finalTextMarkdown.Length + 2 + targetMarkdown.Length <= maxChars;
            IReadOnlyList<string> targetParts = targetMarkdown.Length == 0 || appendTargets
                ? Array.Empty<string>()
                : DocumentReaderEngine.SplitAdapterProjection("\n\n" + targetMarkdown, maxChars);
            int totalParts = parts.Count + targetParts.Count;
            int emittedPart = 0;
            for (int part = 0; part < parts.Count; part++) {
                string markdown = parts[part].Markdown;
                if (part == parts.Count - 1 && appendTargets) markdown += "\n\n" + targetMarkdown;
                yield return new ReaderChunk {
                    Id = totalParts == 1 ? "opml-" + currentSource : "opml-" + currentSource + "-part-" + (emittedPart + 1),
                    Kind = ReaderInputKind.Opml, Text = parts[part].Text,
                    Markdown = markdown,
                    ContinuesPreviousChunk = emittedPart > 0,
                    Location = new ReaderLocation { Path = sourceName, BlockIndex = emittedIndex++, SourceBlockIndex = currentSource,
                        HeadingPath = headingPath, SourceBlockKind = "outline", BlockAnchor = "opml-outline-" + currentSource },
                    Diagnostics = new ReaderChunkDiagnostics { SourceKind = "opml" }, Warnings = TakeWarnings()
                };
                emittedPart++;
            }
            foreach (string targetPart in targetParts) {
                yield return new ReaderChunk {
                    Id = "opml-" + currentSource + "-part-" + (emittedPart + 1),
                    Kind = ReaderInputKind.Opml, Text = string.Empty,
                    Markdown = targetPart,
                    ContinuesPreviousChunk = true,
                    Location = new ReaderLocation { Path = sourceName, BlockIndex = emittedIndex++, SourceBlockIndex = currentSource,
                        HeadingPath = headingPath, SourceBlockKind = "outline-target", BlockAnchor = "opml-outline-" + currentSource },
                    Diagnostics = new ReaderChunkDiagnostics { SourceKind = "opml" }, Warnings = TakeWarnings()
                };
                emittedPart++;
            }
            foreach (OpmlOutline child in outline.Children) foreach (ReaderChunk chunk in BuildOutline(child, level + 1, headingPath)) yield return chunk;
        }
    }

    private static IReadOnlyList<HeadingProjectionPart> SplitHeadingProjection(string text, string prefix, int maxChars) {
        int effectiveMaxChars = Math.Max(1, maxChars);
        IReadOnlyList<string> prefixParts = DocumentReaderEngine.SplitAdapterProjection(prefix, effectiveMaxChars);
        var parts = new List<HeadingProjectionPart>();
        for (int index = 0; index + 1 < prefixParts.Count; index++) {
            parts.Add(new HeadingProjectionPart(string.Empty, prefixParts[index]));
        }
        string finalPrefix = prefixParts.Count == 0 ? string.Empty : prefixParts[prefixParts.Count - 1];
        int firstTextBudget = finalPrefix.Length >= effectiveMaxChars
            ? effectiveMaxChars : effectiveMaxChars - finalPrefix.Length;
        IReadOnlyList<string> textParts = text.Length == 0
            ? Array.Empty<string>()
            : DocumentReaderEngine.SplitAdapterProjection(text, firstTextBudget, effectiveMaxChars);
        if (finalPrefix.Length >= effectiveMaxChars) {
            parts.Add(new HeadingProjectionPart(string.Empty, finalPrefix));
            foreach (string textPart in textParts) parts.Add(new HeadingProjectionPart(textPart, textPart));
        } else if (textParts.Count == 0) {
            parts.Add(new HeadingProjectionPart(string.Empty, finalPrefix));
        } else {
            parts.Add(new HeadingProjectionPart(textParts[0], finalPrefix + textParts[0]));
            for (int index = 1; index < textParts.Count; index++) {
                parts.Add(new HeadingProjectionPart(textParts[index], textParts[index]));
            }
        }
        return parts;
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

    private sealed class HeadingProjectionPart {
        internal HeadingProjectionPart(string text, string markdown) { Text = text; Markdown = markdown; }
        internal string Text { get; }
        internal string Markdown { get; }
    }

    private static void ApplyReaderLimit(OpmlReadOptions options, long? maxBytes) {
        if (maxBytes.HasValue) options.MaxInputBytes = Math.Min(options.MaxInputBytes, maxBytes.Value);
    }

}
