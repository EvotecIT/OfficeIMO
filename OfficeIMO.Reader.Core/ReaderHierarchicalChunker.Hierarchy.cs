using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace OfficeIMO.Reader;

public static partial class ReaderHierarchicalChunker {
    private static IReadOnlyList<ReaderChunkHierarchyNode> BuildHierarchy(
        ChunkingState state,
        OfficeDocumentSource source) {
        string rawRootTitle = BuildRootTitle(source);
        string rootTitle = LimitHierarchyValue(rawRootTitle, state);
        string rootIdentity = !string.IsNullOrWhiteSpace(source.SourceId)
            ? "id:" + source.SourceId!
            : !string.IsNullOrWhiteSpace(source.Path)
                ? "path:" + source.Path!
                : "title:" + rawRootTitle;
        var root = new MutableHierarchyNode(
            "node:" + ComputeSha256Hex("document|" + rootIdentity),
            parentNodeId: null,
            ReaderChunkHierarchyNodeKind.Document,
            rootTitle,
            depth: 0,
            rootTitle);
        var nodes = new List<MutableHierarchyNode> { root };
        var nodesById = new Dictionary<string, MutableHierarchyNode>(StringComparer.Ordinal) { [root.Id] = root };
        var reusable = new Dictionary<string, MutableHierarchyNode>(StringComparer.Ordinal);

        for (int index = 0; index < state.Chunks.Count; index++) {
            state.CancellationToken.ThrowIfCancellationRequested();
            ReaderChunk chunk = state.Chunks[index];
            ReaderLocation location = chunk.Location ?? new ReaderLocation();
            MutableHierarchyNode parent = root;

            string? containerKey = BuildContainerKey(location, state, out string? containerTitle);
            if (containerKey != null && containerTitle != null && parent.Depth < state.Options.MaxHierarchyDepth) {
                parent = GetOrCreateNode(
                    parent,
                    ReaderChunkHierarchyNodeKind.Container,
                    containerKey,
                    containerTitle,
                    nodes,
                    nodesById,
                    reusable,
                    state);
            }

            string? hierarchyHeadingPath = ReaderHeadingPath.GetValidatedHierarchyPath(location);
            HierarchyHeading[] headings = SplitHeadingPath(hierarchyHeadingPath ?? location.HeadingPath, state);
            state.HeadingSlugsByChunkId.TryGetValue(chunk.Id, out IReadOnlyList<string?>? fallbackHeadingSlugs);
            IReadOnlyList<string?>? headingSlugs = fallbackHeadingSlugs != null && fallbackHeadingSlugs.Count == headings.Length
                ? fallbackHeadingSlugs
                : null;
            int remainingDepth = Math.Max(0, state.Options.MaxHierarchyDepth - parent.Depth);
            if (headings.Length > remainingDepth) {
                state.AddLimitDiagnostic("hierarchical-depth-limit", state.Options.MaxHierarchyDepth, "hierarchy depth");
                if (headingSlugs != null) headingSlugs = CollapseHeadingSlugs(headingSlugs, remainingDepth);
                headings = CollapseHeadingDepth(headings, remainingDepth, state);
            }
            for (int headingIndex = 0; headingIndex < headings.Length; headingIndex++) {
                string? headingSlug = headingSlugs != null && headingSlugs.Count == headings.Length
                    ? headingSlugs[headingIndex]
                    : headingIndex == headings.Length - 1
                        ? location.HeadingSlug
                        : null;
                string key = BuildHierarchyIdentity(
                    "heading",
                    headings[headingIndex].Identity,
                    string.IsNullOrWhiteSpace(headingSlug) ? null : headingSlug!.Trim());
                parent = GetOrCreateNode(
                    parent,
                    ReaderChunkHierarchyNodeKind.Heading,
                    key,
                    headings[headingIndex].Title,
                    nodes,
                    nodesById,
                    reusable,
                    state);
            }

            string rawLeafTitle = !string.IsNullOrWhiteSpace(location.BlockAnchor)
                ? location.BlockAnchor!
                : "Chunk " + (index + 1).ToString(CultureInfo.InvariantCulture);
            string leafTitle = LimitHierarchyValue(rawLeafTitle, state);
            var leaf = new MutableHierarchyNode(
                "node:" + ComputeSha256Hex(parent.Id + "|chunk|" + chunk.Id),
                parent.Id,
                ReaderChunkHierarchyNodeKind.Chunk,
                leafTitle,
                parent.Depth + 1,
                LimitHierarchyValue(parent.Path + " > " + leafTitle, state)) {
                ChunkId = chunk.Id,
                TokenCount = chunk.TokenEstimate ?? state.Segments[index].TokenCount
            };
            nodes.Add(leaf);
            nodesById.Add(leaf.Id, leaf);
            parent.ChildNodeIds.Add(leaf.Id);
            AddTokensToAncestors(parent, leaf.TokenCount, nodesById);
        }

        var result = new ReaderChunkHierarchyNode[nodes.Count];
        for (int index = 0; index < nodes.Count; index++) result[index] = nodes[index].Freeze();
        return result;
    }

    private static MutableHierarchyNode GetOrCreateNode(
        MutableHierarchyNode parent,
        ReaderChunkHierarchyNodeKind kind,
        string key,
        string title,
        ICollection<MutableHierarchyNode> nodes,
        IDictionary<string, MutableHierarchyNode> nodesById,
        IDictionary<string, MutableHierarchyNode> reusable,
        ChunkingState state) {
        string lookup = BuildHierarchyIdentity(parent.Id, kind.ToString(), key);
        if (reusable.TryGetValue(lookup, out MutableHierarchyNode? existing)) return existing;

        var created = new MutableHierarchyNode(
            "node:" + ComputeSha256Hex(lookup),
            parent.Id,
            kind,
            title,
            parent.Depth + 1,
            LimitHierarchyValue(parent.Path + " > " + title, state));
        reusable.Add(lookup, created);
        nodes.Add(created);
        nodesById.Add(created.Id, created);
        parent.ChildNodeIds.Add(created.Id);
        return created;
    }

    private static string BuildHierarchyIdentity(params string?[] values) {
        var builder = new System.Text.StringBuilder();
        foreach (string? value in values) AppendSegmentIdentity(builder, value);
        return builder.ToString();
    }

    private static void AddTokensToAncestors(
        MutableHierarchyNode node,
        long tokenCount,
        IReadOnlyDictionary<string, MutableHierarchyNode> nodesById) {
        MutableHierarchyNode? current = node;
        while (current != null) {
            current.TokenCount += tokenCount;
            if (current.ParentNodeId == null || !nodesById.TryGetValue(current.ParentNodeId, out current)) break;
        }
    }

    private static string BuildRootTitle(OfficeDocumentSource source) {
        if (!string.IsNullOrWhiteSpace(source.Title)) return source.Title!.Trim();
        if (!string.IsNullOrWhiteSpace(source.Path)) {
            string name = ReaderLogicalPath.GetFileName(source.Path!);
            if (!string.IsNullOrWhiteSpace(name)) return name;
        }
        return "Document";
    }

    private static string? BuildContainerKey(ReaderLocation location, ChunkingState state, out string? title) {
        if (!string.IsNullOrWhiteSpace(location.Sheet)) {
            string sheet = location.Sheet!.Trim();
            title = LimitHierarchyValue("Sheet: " + sheet, state);
            return "sheet:" + sheet;
        }
        if (location.Slide.HasValue) {
            title = LimitHierarchyValue("Slide " + location.Slide.Value.ToString(CultureInfo.InvariantCulture), state);
            return "slide:" + location.Slide.Value.ToString(CultureInfo.InvariantCulture);
        }
        if (location.Page.HasValue) {
            title = LimitHierarchyValue("Page " + location.Page.Value.ToString(CultureInfo.InvariantCulture), state);
            return "page:" + location.Page.Value.ToString(CultureInfo.InvariantCulture);
        }
        title = null;
        return null;
    }

    private static HierarchyHeading[] SplitHeadingPath(string? headingPath, ChunkingState state) {
        IReadOnlyList<string> candidates = ReaderHeadingPath.Split(headingPath);
        if (candidates.Count == 0) return Array.Empty<HierarchyHeading>();
        var headings = new List<HierarchyHeading>(candidates.Count);
        for (int index = 0; index < candidates.Count; index++) {
            string heading = candidates[index].Trim();
            if (heading.Length > 0) headings.Add(new HierarchyHeading(
                heading,
                LimitHierarchyValue(heading, state)));
        }
        return headings.Count == 0 ? Array.Empty<HierarchyHeading>() : headings.ToArray();
    }

    private static string LimitHierarchyValue(string value, ChunkingState state) {
        if (value.Length <= state.Options.MaxContextCharacters) return value;
        state.AddLimitDiagnostic(
            "hierarchical-context-character-limit",
            state.Options.MaxContextCharacters,
            "context characters");
        return TruncateAtCharacterBoundary(value, state.Options.MaxContextCharacters);
    }

    private static HierarchyHeading[] CollapseHeadingDepth(
        HierarchyHeading[] headings,
        int availableDepth,
        ChunkingState state) {
        if (availableDepth <= 0) return Array.Empty<HierarchyHeading>();
        if (headings.Length <= availableDepth) return headings;
        var collapsed = new HierarchyHeading[availableDepth];
        for (int index = 0; index < availableDepth - 1; index++) collapsed[index] = headings[index];
        int collapsedCount = headings.Length - availableDepth + 1;
        var identities = new string[collapsedCount];
        var titles = new string[collapsedCount];
        for (int index = 0; index < collapsedCount; index++) {
            identities[index] = headings[availableDepth - 1 + index].Identity;
            titles[index] = headings[availableDepth - 1 + index].Title;
        }
        collapsed[availableDepth - 1] = new HierarchyHeading(
            BuildHierarchyIdentity(identities),
            LimitHierarchyValue(string.Join(" > ", titles), state));
        return collapsed;
    }

    private static IReadOnlyList<string?> CollapseHeadingSlugs(
        IReadOnlyList<string?> slugs,
        int availableDepth) {
        if (availableDepth <= 0) return Array.Empty<string?>();
        if (slugs.Count <= availableDepth) return slugs;
        var collapsed = new string?[availableDepth];
        for (int index = 0; index < availableDepth - 1; index++) collapsed[index] = slugs[index];
        var trailing = new string?[slugs.Count - availableDepth + 1];
        for (int index = 0; index < trailing.Length; index++) trailing[index] = slugs[availableDepth - 1 + index];
        collapsed[availableDepth - 1] = BuildHierarchyIdentity(trailing);
        return collapsed;
    }

    private readonly struct HierarchyHeading {
        internal HierarchyHeading(string identity, string title) {
            Identity = identity;
            Title = title;
        }

        internal string Identity { get; }
        internal string Title { get; }
    }

    private sealed class MutableHierarchyNode {
        internal MutableHierarchyNode(
            string id,
            string? parentNodeId,
            ReaderChunkHierarchyNodeKind kind,
            string title,
            int depth,
            string path) {
            Id = id;
            ParentNodeId = parentNodeId;
            Kind = kind;
            Title = title;
            Depth = depth;
            Path = path;
        }

        internal string Id { get; }
        internal string? ParentNodeId { get; }
        internal ReaderChunkHierarchyNodeKind Kind { get; }
        internal string Title { get; }
        internal int Depth { get; }
        internal string Path { get; }
        internal long TokenCount { get; set; }
        internal List<string> ChildNodeIds { get; } = new List<string>();
        internal string? ChunkId { get; set; }

        internal ReaderChunkHierarchyNode Freeze() => new ReaderChunkHierarchyNode {
            Id = Id,
            ParentNodeId = ParentNodeId,
            Kind = Kind,
            Title = Title,
            Depth = Depth,
            Path = Path,
            TokenCount = TokenCount,
            ChildNodeIds = ChildNodeIds.Count == 0 ? Array.Empty<string>() : ChildNodeIds.ToArray(),
            ChunkId = ChunkId
        };
    }
}
