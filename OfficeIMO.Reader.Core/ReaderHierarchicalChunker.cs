using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>Creates bounded token-aware chunks and a document/container/heading hierarchy.</summary>
public static partial class ReaderHierarchicalChunker {
    private static readonly ConditionalWeakTable<ReaderChunk, FallbackHeadingIdentityPath> FallbackHeadingIdentities =
        new ConditionalWeakTable<ReaderChunk, FallbackHeadingIdentityPath>();
    private static readonly ReaderChunk FallbackInputLimitMarker = new ReaderChunk();

    /// <summary>Chunks an already-read rich document.</summary>
    public static ReaderChunkHierarchyResult Chunk(
        OfficeDocumentReadResult document,
        ReaderHierarchicalChunkingOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        ReaderHierarchicalChunkingOptions normalized = Normalize(options);
        return ChunkCore(
            GetSourceChunks(document, normalized.MaxInputChunks, cancellationToken),
            document.Source ?? new OfficeDocumentSource(),
            normalized,
            cancellationToken,
            enforceSingleSourceIdentity: false);
    }

    /// <summary>Chunks an existing source-ordered collection.</summary>
    public static ReaderChunkHierarchyResult Chunk(
        IEnumerable<ReaderChunk> chunks,
        ReaderHierarchicalChunkingOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (chunks == null) throw new ArgumentNullException(nameof(chunks));
        return ChunkCore(
            chunks,
            new OfficeDocumentSource(),
            Normalize(options),
            cancellationToken,
            enforceSingleSourceIdentity: true);
    }

    private static ReaderChunkHierarchyResult ChunkCore(
        IEnumerable<ReaderChunk> source,
        OfficeDocumentSource sourceInfo,
        ReaderHierarchicalChunkingOptions options,
        CancellationToken cancellationToken,
        bool enforceSingleSourceIdentity) {
        var state = new ChunkingState(options, cancellationToken, enforceSingleSourceIdentity);
        int inputIndex = 0;
        using IEnumerator<ReaderChunk> enumerator = source.GetEnumerator();
        while (!state.OutputLimitReached) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!enumerator.MoveNext()) break;
            ReaderChunk? chunk = enumerator.Current;
            if (ReferenceEquals(chunk, FallbackInputLimitMarker)) {
                state.AddLimitDiagnostic("hierarchical-input-chunk-limit", options.MaxInputChunks, "input chunks");
                break;
            }
            if (inputIndex >= options.MaxInputChunks) {
                state.AddLimitDiagnostic("hierarchical-input-chunk-limit", options.MaxInputChunks, "input chunks");
                break;
            }
            if (chunk == null) {
                state.AddDiagnostic("hierarchical-null-input-chunk", "A null input chunk was skipped.");
                inputIndex++;
                continue;
            }
            state.AddSourceChunk(chunk, inputIndex);
            inputIndex++;
        }

        OfficeDocumentSource effectiveSource = ResolveSource(sourceInfo, state.Chunks);
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

    private static IEnumerable<ReaderChunk> GetSourceChunks(
        OfficeDocumentReadResult document,
        int maximumInputChunks,
        CancellationToken cancellationToken) {
        IReadOnlyList<ReaderChunk> chunks = document.Chunks ?? Array.Empty<ReaderChunk>();
        if (chunks.Count > 0) {
            for (int chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++) {
                ReaderChunk? chunk = chunks[chunkIndex];
                yield return chunk == null ? null! : InheritDocumentSource(chunk, document.Source);
            }
            yield break;
        }

        var headings = new List<FallbackHeading>();
        int index = 0;
        foreach (FallbackBlock fallback in EnumerateFallbackBlocks(document, maximumInputChunks, cancellationToken)) {
            if (fallback.IsLimitMarker) {
                yield return FallbackInputLimitMarker;
                yield break;
            }
            OfficeDocumentBlock block = fallback.Block;
            string text = block.Text ?? string.Empty;
            bool isHeading = string.Equals(block.Kind?.Trim(), "heading", StringComparison.OrdinalIgnoreCase);
            ReaderLocation location = CloneLocation(block.Location);
            InheritPageLocation(location, fallback.Page);
            location.Path ??= document.Source?.Path;
            location.SourceBlockIndex ??= index;
            location.SourceBlockKind ??= block.Kind;
            location.BlockAnchor ??= string.IsNullOrWhiteSpace(block.Id) ? "block-" + index.ToString(CultureInfo.InvariantCulture) : block.Id;
            if (isHeading) UpdateFallbackHeadings(headings, block.Level ?? 1, text, location.BlockAnchor);
            if (string.IsNullOrWhiteSpace(location.HeadingPath)) {
                location.HeadingPath = BuildFallbackHeadingDisplayPath(headings);
                ReaderHeadingPath.SetHierarchyPath(location, BuildFallbackHierarchyHeadingPath(headings));
                location.HeadingSlug ??= BuildFallbackHeadingSlug(headings);
            } else if (isHeading) {
                location.HeadingSlug ??= location.BlockAnchor;
            }
            string sourceIdentity = document.Source?.SourceId ?? document.Source?.Path ?? string.Empty;
            var sourceChunk = new ReaderChunk {
                Id = "block:" + ComputeSha256Hex(sourceIdentity + "|" + block.Id + "|" + index.ToString(CultureInfo.InvariantCulture)),
                Kind = document.Kind,
                Location = location,
                SourceId = document.Source?.SourceId,
                SourceHash = document.Source?.SourceHash,
                SourceLastWriteUtc = document.Source?.LastWriteUtc,
                SourceLengthBytes = document.Source?.LengthBytes,
                Text = text
            };
            FallbackHeadingIdentities.Add(
                sourceChunk,
                new FallbackHeadingIdentityPath(headings.Select(heading => heading.Slug).ToArray()));
            yield return sourceChunk;
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

    private static IEnumerable<FallbackBlock> EnumerateFallbackBlocks(
        OfficeDocumentReadResult document,
        int maximumInputChunks,
        CancellationToken cancellationToken) {
        var seen = new HashSet<OfficeDocumentBlock>(ReferenceIdentityComparer<OfficeDocumentBlock>.Instance);
        var seenIds = new HashSet<string>(StringComparer.Ordinal);
        IReadOnlyList<OfficeDocumentPage> pages = document.Pages ?? Array.Empty<OfficeDocumentPage>();
        int maximumInspections = (int)Math.Min(int.MaxValue, (long)maximumInputChunks * 4L);
        int emittedBlocks = 0;
        IReadOnlyList<OfficeDocumentBlock> documentBlocks = document.Blocks ?? Array.Empty<OfficeDocumentBlock>();
        int documentInspectionCount = Math.Min(documentBlocks.Count, maximumInspections);
        var retainedDocumentBlocks = new List<OfficeDocumentBlock>(Math.Min(documentInspectionCount, maximumInputChunks));
        bool documentLimitReached = documentBlocks.Count > documentInspectionCount;
        for (int blockIndex = 0; blockIndex < documentInspectionCount; blockIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeDocumentBlock block = documentBlocks[blockIndex];
            if (!TryRegisterFallbackBlock(block, seen, seenIds)) continue;
            if (emittedBlocks >= maximumInputChunks) {
                documentLimitReached = true;
                break;
            }
            retainedDocumentBlocks.Add(block);
            emittedBlocks++;
        }

        PageBlockIndex pageIndex = IndexPageBlocks(pages, retainedDocumentBlocks, cancellationToken);
        for (int blockIndex = 0; blockIndex < retainedDocumentBlocks.Count; blockIndex++) {
            OfficeDocumentBlock block = retainedDocumentBlocks[blockIndex];
            if (!pageIndex.ByReference.TryGetValue(block, out OfficeDocumentPage? page) && !string.IsNullOrWhiteSpace(block.Id)) {
                pageIndex.ById.TryGetValue(block.Id!, out page);
            }
            yield return new FallbackBlock(block, page);
        }
        if (documentLimitReached) {
            yield return FallbackBlock.LimitMarker;
            yield break;
        }

        int inspectedPageBlocks = 0;
        for (int pageIndexValue = 0; pageIndexValue < pages.Count; pageIndexValue++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeDocumentPage page = pages[pageIndexValue];
            if (page?.Blocks == null) continue;
            IReadOnlyList<OfficeDocumentBlock> pageBlocks = page.Blocks;
            for (int blockIndex = 0; blockIndex < pageBlocks.Count; blockIndex++) {
                if (inspectedPageBlocks >= maximumInspections) {
                    yield return FallbackBlock.LimitMarker;
                    yield break;
                }
                cancellationToken.ThrowIfCancellationRequested();
                inspectedPageBlocks++;
                OfficeDocumentBlock block = pageBlocks[blockIndex];
                if (!TryRegisterFallbackBlock(block, seen, seenIds)) continue;
                if (emittedBlocks >= maximumInputChunks) {
                    yield return FallbackBlock.LimitMarker;
                    yield break;
                }
                emittedBlocks++;
                yield return new FallbackBlock(block, page);
            }
        }
    }

    private static bool TryRegisterFallbackBlock(
        OfficeDocumentBlock? block,
        ISet<OfficeDocumentBlock> seen,
        ISet<string> seenIds) {
        if (block == null || !seen.Add(block)) return false;
        return string.IsNullOrWhiteSpace(block.Id) || seenIds.Add(block.Id!);
    }

    private static PageBlockIndex IndexPageBlocks(
        IReadOnlyList<OfficeDocumentPage> pages,
        IReadOnlyList<OfficeDocumentBlock> targetBlocks,
        CancellationToken cancellationToken) {
        var byReference = new Dictionary<OfficeDocumentBlock, OfficeDocumentPage>(
            ReferenceIdentityComparer<OfficeDocumentBlock>.Instance);
        var byId = new Dictionary<string, OfficeDocumentPage>(StringComparer.Ordinal);
        var remaining = new HashSet<OfficeDocumentBlock>(
            targetBlocks,
            ReferenceIdentityComparer<OfficeDocumentBlock>.Instance);
        var targetsById = new Dictionary<string, OfficeDocumentBlock>(StringComparer.Ordinal);
        for (int targetIndex = 0; targetIndex < targetBlocks.Count; targetIndex++) {
            OfficeDocumentBlock target = targetBlocks[targetIndex];
            if (!string.IsNullOrWhiteSpace(target.Id)) {
                targetsById[target.Id!] = target;
            }
        }

        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeDocumentPage page = pages[pageIndex];
            if (page?.Blocks == null) continue;
            foreach (OfficeDocumentBlock block in page.Blocks) {
                cancellationToken.ThrowIfCancellationRequested();
                if (block == null) continue;
                if (remaining.Contains(block)) {
                    byReference[block] = page;
                    remaining.Remove(block);
                }
                if (!string.IsNullOrWhiteSpace(block.Id) &&
                    targetsById.TryGetValue(block.Id!, out OfficeDocumentBlock? target)) {
                    byId[block.Id!] = page;
                    remaining.Remove(target);
                }
                if (remaining.Count == 0) return new PageBlockIndex(byReference, byId);
            }
        }
        return new PageBlockIndex(byReference, byId);
    }

    private sealed class PageBlockIndex {
        internal PageBlockIndex(
            IReadOnlyDictionary<OfficeDocumentBlock, OfficeDocumentPage> byReference,
            IReadOnlyDictionary<string, OfficeDocumentPage> byId) {
            ByReference = byReference;
            ById = byId;
        }

        internal IReadOnlyDictionary<OfficeDocumentBlock, OfficeDocumentPage> ByReference { get; }
        internal IReadOnlyDictionary<string, OfficeDocumentPage> ById { get; }
    }

    private static ReaderChunk InheritDocumentSource(ReaderChunk chunk, OfficeDocumentSource? source) {
        if (source == null) return chunk;
        bool inheritSourceId = string.IsNullOrWhiteSpace(chunk.SourceId) && !string.IsNullOrWhiteSpace(source.SourceId);
        bool inheritPath = string.IsNullOrWhiteSpace(chunk.Location?.Path) && !string.IsNullOrWhiteSpace(source.Path);
        bool inheritSourceHash = string.IsNullOrWhiteSpace(chunk.SourceHash) && !string.IsNullOrWhiteSpace(source.SourceHash);
        bool inheritLastWriteUtc = !chunk.SourceLastWriteUtc.HasValue && source.LastWriteUtc.HasValue;
        bool inheritLengthBytes = !chunk.SourceLengthBytes.HasValue && source.LengthBytes.HasValue;
        if (!inheritSourceId && !inheritPath && !inheritSourceHash && !inheritLastWriteUtc && !inheritLengthBytes) {
            return chunk;
        }

        ReaderLocation location = CloneLocation(chunk.Location);
        if (inheritPath) location.Path = source.Path;
        return new ReaderChunk {
            Id = chunk.Id,
            Kind = chunk.Kind,
            Location = location,
            SourceId = inheritSourceId ? source.SourceId : chunk.SourceId,
            SourceHash = inheritSourceHash ? source.SourceHash : chunk.SourceHash,
            ChunkHash = chunk.ChunkHash,
            SourceLastWriteUtc = inheritLastWriteUtc ? source.LastWriteUtc : chunk.SourceLastWriteUtc,
            SourceLengthBytes = inheritLengthBytes ? source.LengthBytes : chunk.SourceLengthBytes,
            TokenEstimate = chunk.TokenEstimate,
            Text = chunk.Text,
            Markdown = chunk.Markdown,
            ContinuesPreviousChunk = chunk.ContinuesPreviousChunk,
            Tables = chunk.Tables,
            Visuals = chunk.Visuals,
            FormFields = chunk.FormFields,
            Actions = chunk.Actions,
            Diagnostics = chunk.Diagnostics,
            Warnings = chunk.Warnings
        };
    }

    private static void InheritPageLocation(ReaderLocation location, OfficeDocumentPage? page) {
        if (page == null) return;
        ReaderLocation container = page.Location ?? new ReaderLocation();
        if (string.IsNullOrWhiteSpace(location.Path)) location.Path = container.Path;
        if (string.IsNullOrWhiteSpace(location.Sheet)) location.Sheet = container.Sheet;
        if (string.IsNullOrWhiteSpace(location.A1Range)) location.A1Range = container.A1Range;
        location.Slide ??= container.Slide;
        string? containerKind = container.SourceBlockKind?.Trim();
        if (string.Equals(containerKind, "sheet", StringComparison.OrdinalIgnoreCase)) {
            if (string.IsNullOrWhiteSpace(location.Sheet)) {
                location.Sheet = !string.IsNullOrWhiteSpace(page.Name)
                    ? page.Name
                    : page.Number > 0
                        ? "Sheet " + page.Number.Value.ToString(CultureInfo.InvariantCulture)
                        : null;
            }
        }
        if (!location.Page.HasValue &&
            !location.Slide.HasValue &&
            string.IsNullOrWhiteSpace(location.Sheet)) {
            int? number = page.Number > 0 ? page.Number : container.Page;
            if (string.Equals(containerKind, "slide", StringComparison.OrdinalIgnoreCase)) {
                location.Slide = number;
            } else {
                location.Page = number;
            }
        }
    }

    private static void UpdateFallbackHeadings(
        List<FallbackHeading> headings,
        int level,
        string text,
        string? slug) {
        int effectiveLevel = Math.Max(1, level);
        for (int index = headings.Count - 1; index >= 0; index--) {
            if (headings[index].Level >= effectiveLevel) headings.RemoveAt(index);
        }
        string title = string.IsNullOrWhiteSpace(text)
            ? "Heading " + effectiveLevel.ToString(CultureInfo.InvariantCulture)
            : text.Trim();
        headings.Add(new FallbackHeading(effectiveLevel, title, slug));
    }

    private static string? BuildFallbackHierarchyHeadingPath(IReadOnlyList<FallbackHeading> headings) {
        return ReaderHeadingPath.Combine(headings.Select(heading => heading.Title));
    }

    private static string? BuildFallbackHeadingDisplayPath(IReadOnlyList<FallbackHeading> headings) {
        string[] titles = headings
            .Select(heading => heading.Title)
            .Where(title => !string.IsNullOrWhiteSpace(title))
            .Select(title => title.Trim())
            .ToArray();
        return titles.Length == 0 ? null : string.Join(" > ", titles);
    }

    private static string? BuildFallbackHeadingSlug(IReadOnlyList<FallbackHeading> headings) =>
        headings.Count == 0 ? null : headings[headings.Count - 1].Slug;

    private static OfficeDocumentSource ResolveSource(
        OfficeDocumentSource source,
        IReadOnlyList<ReaderChunk> chunks) {
        OfficeDocumentSource inferred = InferSource(chunks);
        return new OfficeDocumentSource {
            Path = PreferSourceValue(source.Path, inferred.Path),
            SourceId = PreferSourceValue(source.SourceId, inferred.SourceId),
            SourceHash = PreferSourceValue(source.SourceHash, inferred.SourceHash),
            LastWriteUtc = source.LastWriteUtc ?? inferred.LastWriteUtc,
            LengthBytes = source.LengthBytes ?? inferred.LengthBytes,
            Title = source.Title,
            Author = source.Author,
            Subject = source.Subject,
            Keywords = source.Keywords
        };
    }

    private static string? PreferSourceValue(string? preferred, string? fallback) =>
        !string.IsNullOrWhiteSpace(preferred) ? preferred : fallback;

    private static OfficeDocumentSource InferSource(IReadOnlyList<ReaderChunk> chunks) {
        if (chunks.Count == 0) return new OfficeDocumentSource();
        ReaderChunk first = chunks.FirstOrDefault(chunk => GetSourceIdentity(chunk) != null) ?? chunks[0];
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

        internal ChunkingState(
            ReaderHierarchicalChunkingOptions options,
            CancellationToken cancellationToken,
            bool enforceSingleSourceIdentity) {
            Options = options;
            CancellationToken = cancellationToken;
            TokenCounterId = options.TokenCounter.Id;
            EnforceSingleSourceIdentity = enforceSingleSourceIdentity;
        }

        internal ReaderHierarchicalChunkingOptions Options { get; }
        internal CancellationToken CancellationToken { get; }
        internal string TokenCounterId { get; }
        internal bool EnforceSingleSourceIdentity { get; }
        internal List<ReaderChunk> Chunks { get; } = new List<ReaderChunk>();
        internal List<ReaderChunkSegment> Segments { get; } = new List<ReaderChunkSegment>();
        internal List<OfficeDocumentDiagnostic> Diagnostics { get; } = new List<OfficeDocumentDiagnostic>();
        internal Dictionary<string, IReadOnlyList<string?>> HeadingSlugsByChunkId { get; } =
            new Dictionary<string, IReadOnlyList<string?>>(StringComparer.Ordinal);
        internal long SourceTokenCount { get; set; }
        internal long OutputTokenCount { get; set; }
        internal long OverlapTokenCount { get; set; }
        internal long ContextTokenCount { get; set; }
        internal bool OutputLimitReached { get; set; }
        internal string? SourceIdentity { get; set; }

        internal void AddSourceChunk(ReaderChunk source, int inputIndex) {
            string? sourceIdentity = GetSourceIdentity(source);
            if (SourceIdentity == null) SourceIdentity = sourceIdentity;
            else if (EnforceSingleSourceIdentity && sourceIdentity != null &&
                     !string.Equals(SourceIdentity, sourceIdentity, StringComparison.Ordinal)) {
                throw new InvalidOperationException("One chunk hierarchy cannot contain chunks from multiple source documents.");
            }
            int firstOutputIndex = Chunks.Count;
            SplitSourceChunk(source, inputIndex, this);
            if (FallbackHeadingIdentities.TryGetValue(source, out FallbackHeadingIdentityPath? identityPath)) {
                for (int index = firstOutputIndex; index < Chunks.Count; index++) {
                    HeadingSlugsByChunkId[Chunks[index].Id] = identityPath.Slugs;
                }
            }
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

    private readonly struct FallbackBlock {
        internal FallbackBlock(OfficeDocumentBlock block, OfficeDocumentPage? page) {
            Block = block;
            Page = page;
            IsLimitMarker = false;
        }

        private FallbackBlock(bool isLimitMarker) {
            Block = null!;
            Page = null;
            IsLimitMarker = isLimitMarker;
        }

        internal static FallbackBlock LimitMarker { get; } = new FallbackBlock(true);

        internal OfficeDocumentBlock Block { get; }
        internal OfficeDocumentPage? Page { get; }
        internal bool IsLimitMarker { get; }
    }

    private readonly struct FallbackHeading {
        internal FallbackHeading(int level, string title, string? slug) {
            Level = level;
            Title = title;
            Slug = slug;
        }

        internal int Level { get; }
        internal string Title { get; }
        internal string? Slug { get; }
    }

    private sealed class FallbackHeadingIdentityPath {
        internal FallbackHeadingIdentityPath(IReadOnlyList<string?> slugs) {
            Slugs = slugs;
        }

        internal IReadOnlyList<string?> Slugs { get; }
    }
}
