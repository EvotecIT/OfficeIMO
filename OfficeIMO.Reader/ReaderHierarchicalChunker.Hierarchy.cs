using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace OfficeIMO.Reader;

public static partial class ReaderHierarchicalChunker {
    private static IReadOnlyList<ReaderChunkHierarchyNode> BuildHierarchy(
        ChunkingState state,
        OfficeDocumentSource source) {
        string rootTitle = LimitHierarchyValue(BuildRootTitle(source), state);
        var root = new MutableHierarchyNode(
            "node:" + ComputeSha256Hex("document|" + (source.SourceId ?? source.Path ?? rootTitle)),
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
                    reusable);
            }

            string[] headings = SplitHeadingPath(location.HeadingPath, state);
            int remainingDepth = Math.Max(0, state.Options.MaxHierarchyDepth - parent.Depth);
            if (headings.Length > remainingDepth) {
                state.AddLimitDiagnostic("hierarchical-depth-limit", state.Options.MaxHierarchyDepth, "hierarchy depth");
                headings = CollapseHeadingDepth(headings, remainingDepth);
            }
            for (int headingIndex = 0; headingIndex < headings.Length; headingIndex++) {
                bool finalHeading = headingIndex == headings.Length - 1;
                string key = "heading:" + headings[headingIndex];
                if (finalHeading && !string.IsNullOrWhiteSpace(location.HeadingSlug)) {
                    key += "|slug:" + location.HeadingSlug!.Trim();
                }
                parent = GetOrCreateNode(
                    parent,
                    ReaderChunkHierarchyNodeKind.Heading,
                    key,
                    headings[headingIndex],
                    nodes,
                    nodesById,
                    reusable);
            }

            string leafTitle = !string.IsNullOrWhiteSpace(location.BlockAnchor)
                ? location.BlockAnchor!
                : "Chunk " + (index + 1).ToString(CultureInfo.InvariantCulture);
            var leaf = new MutableHierarchyNode(
                "node:" + ComputeSha256Hex(parent.Id + "|chunk|" + chunk.Id),
                parent.Id,
                ReaderChunkHierarchyNodeKind.Chunk,
                leafTitle,
                parent.Depth + 1,
                parent.Path + " > " + leafTitle) {
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
        IDictionary<string, MutableHierarchyNode> reusable) {
        string lookup = parent.Id + "|" + kind.ToString() + "|" + key;
        if (reusable.TryGetValue(lookup, out MutableHierarchyNode? existing)) return existing;

        var created = new MutableHierarchyNode(
            "node:" + ComputeSha256Hex(lookup),
            parent.Id,
            kind,
            title,
            parent.Depth + 1,
            parent.Path + " > " + title);
        reusable.Add(lookup, created);
        nodes.Add(created);
        nodesById.Add(created.Id, created);
        parent.ChildNodeIds.Add(created.Id);
        return created;
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
            string name = Path.GetFileName(source.Path);
            if (!string.IsNullOrWhiteSpace(name)) return name;
        }
        return "Document";
    }

    private static string? BuildContainerKey(ReaderLocation location, ChunkingState state, out string? title) {
        if (!string.IsNullOrWhiteSpace(location.Sheet)) {
            string sheet = LimitHierarchyValue(location.Sheet!.Trim(), state);
            title = "Sheet: " + sheet;
            return "sheet:" + sheet;
        }
        if (location.Slide.HasValue) {
            title = "Slide " + location.Slide.Value.ToString(CultureInfo.InvariantCulture);
            return "slide:" + location.Slide.Value.ToString(CultureInfo.InvariantCulture);
        }
        if (location.Page.HasValue) {
            title = "Page " + location.Page.Value.ToString(CultureInfo.InvariantCulture);
            return "page:" + location.Page.Value.ToString(CultureInfo.InvariantCulture);
        }
        title = null;
        return null;
    }

    private static string[] SplitHeadingPath(string? headingPath, ChunkingState state) {
        if (string.IsNullOrWhiteSpace(headingPath)) return Array.Empty<string>();
        string normalized = LimitHierarchyValue(headingPath!.Trim(), state);
        string[] candidates = normalized.Split(new[] { " > " }, StringSplitOptions.RemoveEmptyEntries);
        var headings = new List<string>(candidates.Length);
        for (int index = 0; index < candidates.Length; index++) {
            string heading = candidates[index].Trim();
            if (heading.Length > 0) headings.Add(heading);
        }
        return headings.Count == 0 ? Array.Empty<string>() : headings.ToArray();
    }

    private static string LimitHierarchyValue(string value, ChunkingState state) {
        if (value.Length <= state.Options.MaxContextCharacters) return value;
        state.AddLimitDiagnostic(
            "hierarchical-context-character-limit",
            state.Options.MaxContextCharacters,
            "context characters");
        return TruncateAtCharacterBoundary(value, state.Options.MaxContextCharacters);
    }

    private static string[] CollapseHeadingDepth(string[] headings, int availableDepth) {
        if (availableDepth <= 0) return Array.Empty<string>();
        if (headings.Length <= availableDepth) return headings;
        var collapsed = new string[availableDepth];
        for (int index = 0; index < availableDepth - 1; index++) collapsed[index] = headings[index];
        collapsed[availableDepth - 1] = string.Join(" > ", headings, availableDepth - 1, headings.Length - availableDepth + 1);
        return collapsed;
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
