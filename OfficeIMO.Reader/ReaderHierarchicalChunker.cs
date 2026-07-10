using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>Creates bounded token-aware chunks and a document/container/heading hierarchy.</summary>
public static partial class ReaderHierarchicalChunker {
    /// <summary>Chunks an already-read rich document.</summary>
    public static ReaderChunkHierarchyResult Chunk(
        OfficeDocumentReadResult document,
        ReaderHierarchicalChunkingOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return ChunkCore(
            GetSourceChunks(document),
            document.Source ?? new OfficeDocumentSource(),
            Normalize(options),
            cancellationToken);
    }

    /// <summary>Chunks an existing source-ordered collection.</summary>
    public static ReaderChunkHierarchyResult Chunk(
        IEnumerable<ReaderChunk> chunks,
        ReaderHierarchicalChunkingOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (chunks == null) throw new ArgumentNullException(nameof(chunks));
        return ChunkCore(chunks, new OfficeDocumentSource(), Normalize(options), cancellationToken);
    }

    private static ReaderChunkHierarchyResult ChunkCore(
        IEnumerable<ReaderChunk> source,
        OfficeDocumentSource sourceInfo,
        ReaderHierarchicalChunkingOptions options,
        CancellationToken cancellationToken) {
        var state = new ChunkingState(options, cancellationToken);
        int inputIndex = 0;
        using IEnumerator<ReaderChunk> enumerator = source.GetEnumerator();
        while (!state.OutputLimitReached) {
            cancellationToken.ThrowIfCancellationRequested();
            if (inputIndex >= options.MaxInputChunks) {
                state.AddLimitDiagnostic("hierarchical-input-chunk-limit", options.MaxInputChunks, "input chunks");
                break;
            }
            if (!enumerator.MoveNext()) break;
            ReaderChunk? chunk = enumerator.Current;
            if (chunk == null) {
                state.AddDiagnostic("hierarchical-null-input-chunk", "A null input chunk was skipped.");
                inputIndex++;
                continue;
            }
            state.AddSourceChunk(chunk, inputIndex);
            inputIndex++;
        }

        OfficeDocumentSource effectiveSource = HasSourceIdentity(sourceInfo)
            ? sourceInfo
            : InferSource(state.Chunks);
        IReadOnlyList<ReaderChunkHierarchyNode> nodes = BuildHierarchy(state, effectiveSource);
        return new ReaderChunkHierarchyResult {
            Source = CloneSource(effectiveSource),
            TokenCounterId = state.TokenCounterId,
            RootNodeId = nodes.Count == 0 ? string.Empty : nodes[0].Id,
            SourceTokenCount = state.SourceTokenCount,
            OutputTokenCount = state.OutputTokenCount,
            OverlapTokenCount = state.OverlapTokenCount,
            ContextTokenCount = state.ContextTokenCount,
            Chunks = state.Chunks.Count == 0 ? Array.Empty<ReaderChunk>() : state.Chunks.ToArray(),
            Segments = state.Segments.Count == 0 ? Array.Empty<ReaderChunkSegment>() : state.Segments.ToArray(),
            Nodes = nodes,
            Diagnostics = state.Diagnostics.Count == 0 ? Array.Empty<OfficeDocumentDiagnostic>() : state.Diagnostics.ToArray()
        };
    }

    private static ReaderHierarchicalChunkingOptions Normalize(ReaderHierarchicalChunkingOptions? options) {
        ReaderHierarchicalChunkingOptions effective = (options ?? new ReaderHierarchicalChunkingOptions()).Clone();
        if (effective.MaxTokens <= 0) throw new ArgumentOutOfRangeException(nameof(effective.MaxTokens));
        if (effective.OverlapTokens < 0 || effective.OverlapTokens >= effective.MaxTokens) {
            throw new ArgumentOutOfRangeException(nameof(effective.OverlapTokens), "Overlap must be non-negative and smaller than MaxTokens.");
        }
        if (effective.MaxInputChunks <= 0) throw new ArgumentOutOfRangeException(nameof(effective.MaxInputChunks));
        if (effective.MaxOutputChunks <= 0) throw new ArgumentOutOfRangeException(nameof(effective.MaxOutputChunks));
        if (effective.MaxHierarchyDepth <= 0) throw new ArgumentOutOfRangeException(nameof(effective.MaxHierarchyDepth));
        if (effective.MaxContextCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(effective.MaxContextCharacters));
        if (effective.TokenCounter == null) throw new ArgumentNullException(nameof(effective.TokenCounter));
        string counterId = effective.TokenCounter.Id;
        if (string.IsNullOrWhiteSpace(counterId) || !string.Equals(counterId, counterId.Trim(), StringComparison.Ordinal)) {
            throw new ArgumentException("Token counter id must be non-empty and normalized.", nameof(effective.TokenCounter));
        }
        return effective;
    }

    private static IEnumerable<ReaderChunk> GetSourceChunks(OfficeDocumentReadResult document) {
        IReadOnlyList<ReaderChunk> chunks = document.Chunks ?? Array.Empty<ReaderChunk>();
        if (chunks.Count > 0) {
            for (int chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++) yield return chunks[chunkIndex];
            yield break;
        }

        var headings = new List<KeyValuePair<int, string>>();
        int index = 0;
        foreach (OfficeDocumentBlock block in OfficeDocumentModelTraversal.Blocks(document)) {
            string text = block.Text ?? string.Empty;
            bool isHeading = string.Equals(block.Kind?.Trim(), "heading", StringComparison.OrdinalIgnoreCase);
            if (isHeading) UpdateFallbackHeadings(headings, block.Level ?? 1, text);
            ReaderLocation location = CloneLocation(block.Location);
            location.SourceBlockIndex ??= index;
            location.SourceBlockKind ??= block.Kind;
            location.BlockAnchor ??= string.IsNullOrWhiteSpace(block.Id) ? "block-" + index.ToString(CultureInfo.InvariantCulture) : block.Id;
            location.HeadingPath ??= BuildFallbackHeadingPath(headings);
            if (isHeading) location.HeadingSlug ??= string.IsNullOrWhiteSpace(block.Id) ? null : block.Id;
            yield return new ReaderChunk {
                Id = "block:" + ComputeSha256Hex((document.Source?.SourceId ?? string.Empty) + "|" + block.Id + "|" + index.ToString(CultureInfo.InvariantCulture)),
                Kind = document.Kind,
                Location = location,
                SourceId = document.Source?.SourceId,
                SourceHash = document.Source?.SourceHash,
                SourceLastWriteUtc = document.Source?.LastWriteUtc,
                SourceLengthBytes = document.Source?.LengthBytes,
                Text = text
            };
            index++;
        }
        if (index == 0 && !string.IsNullOrEmpty(document.Markdown)) {
            yield return new ReaderChunk {
                Id = "document:" + ComputeSha256Hex((document.Source?.SourceId ?? document.Source?.Path ?? string.Empty) + "|markdown"),
                Kind = document.Kind,
                Location = new ReaderLocation { Path = document.Source?.Path },
                SourceId = document.Source?.SourceId,
                SourceHash = document.Source?.SourceHash,
                SourceLastWriteUtc = document.Source?.LastWriteUtc,
                SourceLengthBytes = document.Source?.LengthBytes,
                Text = document.Markdown!,
                Markdown = document.Markdown
            };
        }
    }

    private static void UpdateFallbackHeadings(List<KeyValuePair<int, string>> headings, int level, string text) {
        int effectiveLevel = Math.Max(1, level);
        for (int index = headings.Count - 1; index >= 0; index--) {
            if (headings[index].Key >= effectiveLevel) headings.RemoveAt(index);
        }
        string title = string.IsNullOrWhiteSpace(text)
            ? "Heading " + effectiveLevel.ToString(CultureInfo.InvariantCulture)
            : text.Trim();
        headings.Add(new KeyValuePair<int, string>(effectiveLevel, title));
    }

    private static string? BuildFallbackHeadingPath(IReadOnlyList<KeyValuePair<int, string>> headings) {
        if (headings.Count == 0) return null;
        return string.Join(" > ", headings.Select(heading => heading.Value));
    }

    private static bool HasSourceIdentity(OfficeDocumentSource source) =>
        !string.IsNullOrWhiteSpace(source.SourceId) ||
        !string.IsNullOrWhiteSpace(source.Path) ||
        !string.IsNullOrWhiteSpace(source.Title);

    private static OfficeDocumentSource InferSource(IReadOnlyList<ReaderChunk> chunks) {
        if (chunks.Count == 0) return new OfficeDocumentSource();
        ReaderChunk first = chunks[0];
        return new OfficeDocumentSource {
            Path = first.Location?.Path,
            SourceId = first.SourceId,
            SourceHash = first.SourceHash,
            LastWriteUtc = first.SourceLastWriteUtc,
            LengthBytes = first.SourceLengthBytes
        };
    }

    private static string TruncateAtCharacterBoundary(string value, int maximumCharacters) {
        if (value.Length <= maximumCharacters) return value;
        int length = maximumCharacters;
        if (length > 0 &&
            length < value.Length &&
            char.IsHighSurrogate(value[length - 1]) &&
            char.IsLowSurrogate(value[length])) {
            length--;
        }
        return length == 0 ? string.Empty : value.Substring(0, length);
    }

    private static OfficeDocumentSource CloneSource(OfficeDocumentSource source) => new OfficeDocumentSource {
        Path = source.Path,
        SourceId = source.SourceId,
        SourceHash = source.SourceHash,
        LastWriteUtc = source.LastWriteUtc,
        LengthBytes = source.LengthBytes,
        Title = source.Title,
        Author = source.Author,
        Subject = source.Subject,
        Keywords = source.Keywords
    };

    private sealed class ChunkingState {
        private readonly HashSet<string> _diagnosticCodes = new HashSet<string>(StringComparer.Ordinal);

        internal ChunkingState(ReaderHierarchicalChunkingOptions options, CancellationToken cancellationToken) {
            Options = options;
            CancellationToken = cancellationToken;
            TokenCounterId = options.TokenCounter.Id;
        }

        internal ReaderHierarchicalChunkingOptions Options { get; }
        internal CancellationToken CancellationToken { get; }
        internal string TokenCounterId { get; }
        internal List<ReaderChunk> Chunks { get; } = new List<ReaderChunk>();
        internal List<ReaderChunkSegment> Segments { get; } = new List<ReaderChunkSegment>();
        internal List<OfficeDocumentDiagnostic> Diagnostics { get; } = new List<OfficeDocumentDiagnostic>();
        internal long SourceTokenCount { get; set; }
        internal long OutputTokenCount { get; set; }
        internal long OverlapTokenCount { get; set; }
        internal long ContextTokenCount { get; set; }
        internal bool OutputLimitReached { get; set; }
        internal string? SourceIdentity { get; set; }

        internal void AddSourceChunk(ReaderChunk source, int inputIndex) {
            string? sourceIdentity = GetSourceIdentity(source);
            if (SourceIdentity == null) SourceIdentity = sourceIdentity;
            else if (sourceIdentity != null && !string.Equals(SourceIdentity, sourceIdentity, StringComparison.Ordinal)) {
                throw new InvalidOperationException("One chunk hierarchy cannot contain chunks from multiple source documents.");
            }
            SplitSourceChunk(source, inputIndex, this);
        }

        internal void AddLimitDiagnostic(string code, int limit, string target) {
            if (!_diagnosticCodes.Add(code)) return;
            Diagnostics.Add(new OfficeDocumentDiagnostic {
                Severity = OfficeDocumentDiagnosticSeverity.Warning,
                Category = OfficeDocumentDiagnosticCategory.Limit,
                Code = code,
                Message = $"Hierarchical chunking reached the configured {target} limit ({limit.ToString(CultureInfo.InvariantCulture)}).",
                Source = "officeimo.reader.hierarchical-chunker",
                IsRecoverable = true,
                Attributes = new SortedDictionary<string, string>(StringComparer.Ordinal) {
                    ["limit"] = limit.ToString(CultureInfo.InvariantCulture),
                    ["target"] = target
                }
            });
        }

        internal void AddDiagnostic(string code, string message) {
            if (!_diagnosticCodes.Add(code)) return;
            Diagnostics.Add(new OfficeDocumentDiagnostic {
                Severity = OfficeDocumentDiagnosticSeverity.Warning,
                Category = OfficeDocumentDiagnosticCategory.General,
                Code = code,
                Message = message,
                Source = "officeimo.reader.hierarchical-chunker",
                IsRecoverable = true
            });
        }
    }

    private static string? GetSourceIdentity(ReaderChunk source) {
        string? sourcePath = source.Location?.Path;
        return !string.IsNullOrWhiteSpace(source.SourceId)
            ? "id:" + source.SourceId
            : !string.IsNullOrWhiteSpace(sourcePath)
                ? "path:" + sourcePath
                : null;
    }
}
