using System;
using System.Collections.Generic;

namespace OfficeIMO.Reader;

/// <summary>Schema identifiers for the independently versioned hierarchical chunking result.</summary>
public static class ReaderChunkHierarchySchema {
    /// <summary>Schema identifier.</summary>
    public const string Id = "officeimo.reader.chunk-hierarchy";

    /// <summary>Current schema version.</summary>
    public const int Version = 1;
}

/// <summary>Options for bounded, token-aware hierarchical chunking.</summary>
public sealed class ReaderHierarchicalChunkingOptions {
    /// <summary>Maximum tokens in each output chunk, including optional context. Default: 800.</summary>
    public int MaxTokens { get; set; } = 800;

    /// <summary>Maximum content tokens repeated from the previous segment. Default: 80.</summary>
    public int OverlapTokens { get; set; } = 80;

    /// <summary>Maximum source chunks inspected. Default: 10,000.</summary>
    public int MaxInputChunks { get; set; } = 10_000;

    /// <summary>Maximum output chunks emitted. Default: 50,000.</summary>
    public int MaxOutputChunks { get; set; } = 50_000;

    /// <summary>Maximum container and heading hierarchy depth below the document root. Default: 32.</summary>
    public int MaxHierarchyDepth { get; set; } = 32;

    /// <summary>Maximum characters retained in one hierarchy breadcrumb or title. Default: 4,096.</summary>
    public int MaxContextCharacters { get; set; } = 4_096;

    /// <summary>Prefer a source chunk's Markdown over plain text when available. Default: true.</summary>
    public bool PreferMarkdown { get; set; } = true;

    /// <summary>Prefix output text with its container and heading breadcrumb when it fits the token budget. Default: true.</summary>
    public bool IncludeContextInText { get; set; } = true;

    /// <summary>Token counter. Defaults to the dependency-free OfficeIMO heuristic.</summary>
    public IReaderTokenCounter TokenCounter { get; set; } = ReaderHeuristicTokenCounter.Instance;

    internal ReaderHierarchicalChunkingOptions Clone() => new ReaderHierarchicalChunkingOptions {
        MaxTokens = MaxTokens,
        OverlapTokens = OverlapTokens,
        MaxInputChunks = MaxInputChunks,
        MaxOutputChunks = MaxOutputChunks,
        MaxHierarchyDepth = MaxHierarchyDepth,
        MaxContextCharacters = MaxContextCharacters,
        PreferMarkdown = PreferMarkdown,
        IncludeContextInText = IncludeContextInText,
        TokenCounter = TokenCounter
    };
}

/// <summary>Kind of one node in a flattened chunk hierarchy.</summary>
public enum ReaderChunkHierarchyNodeKind {
    /// <summary>Document root.</summary>
    Document = 0,
    /// <summary>Page, slide, or sheet container.</summary>
    Container,
    /// <summary>Logical heading level.</summary>
    Heading,
    /// <summary>Leaf pointing to one output chunk.</summary>
    Chunk
}

/// <summary>One node in a deterministic flattened hierarchy.</summary>
public sealed class ReaderChunkHierarchyNode {
    /// <summary>Deterministic node identifier.</summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>Parent node identifier, or null for the document root.</summary>
    public string? ParentNodeId { get; set; }

    /// <summary>Node kind.</summary>
    public ReaderChunkHierarchyNodeKind Kind { get; set; }

    /// <summary>Human-readable title.</summary>
    public string Title { get; set; } = string.Empty;

    /// <summary>Hierarchy depth, with the root at zero.</summary>
    public int Depth { get; set; }

    /// <summary>Breadcrumb from the root to this node.</summary>
    public string Path { get; set; } = string.Empty;

    /// <summary>Output tokens in this node and all descendants.</summary>
    public long TokenCount { get; set; }

    /// <summary>Child node identifiers in source order.</summary>
    public IReadOnlyList<string> ChildNodeIds { get; set; } = Array.Empty<string>();

    /// <summary>Output chunk identifier for a leaf node.</summary>
    public string? ChunkId { get; set; }
}

/// <summary>Exact source-span and token evidence for one output chunk.</summary>
public sealed class ReaderChunkSegment {
    /// <summary>Output chunk identifier.</summary>
    public string ChunkId { get; set; } = string.Empty;

    /// <summary>Source chunk identifier.</summary>
    public string SourceChunkId { get; set; } = string.Empty;

    /// <summary>Zero-based segment index within the source chunk.</summary>
    public int SegmentIndex { get; set; }

    /// <summary>Inclusive character offset in the selected source representation.</summary>
    public int StartCharacter { get; set; }

    /// <summary>Exclusive character offset in the selected source representation.</summary>
    public int EndCharacter { get; set; }

    /// <summary>Characters repeated from the previous segment.</summary>
    public int OverlapCharacterCount { get; set; }

    /// <summary>Repeated content tokens from the previous segment.</summary>
    public int OverlapTokenCount { get; set; }

    /// <summary>Tokens contributed by source content.</summary>
    public int ContentTokenCount { get; set; }

    /// <summary>Tokens contributed by the optional hierarchy prefix.</summary>
    public int ContextTokenCount { get; set; }

    /// <summary>Total tokens in the output chunk.</summary>
    public int TokenCount { get; set; }

    /// <summary>Container and heading breadcrumb available to retrieval hosts.</summary>
    public string? Context { get; set; }
}

/// <summary>Bounded token-aware chunks plus their deterministic hierarchy and evidence.</summary>
public sealed class ReaderChunkHierarchyResult {
    /// <summary>Hierarchy schema identifier.</summary>
    public string SchemaId { get; set; } = ReaderChunkHierarchySchema.Id;

    /// <summary>Hierarchy schema version.</summary>
    public int SchemaVersion { get; set; } = ReaderChunkHierarchySchema.Version;

    /// <summary>Source document metadata.</summary>
    public OfficeDocumentSource Source { get; set; } = new OfficeDocumentSource();

    /// <summary>Token counter identifier.</summary>
    public string TokenCounterId { get; set; } = string.Empty;

    /// <summary>Document root node identifier.</summary>
    public string RootNodeId { get; set; } = string.Empty;

    /// <summary>Original source tokens before overlap and hierarchy context.</summary>
    public long SourceTokenCount { get; set; }

    /// <summary>Total tokens emitted across output chunks.</summary>
    public long OutputTokenCount { get; set; }

    /// <summary>Total repeated overlap tokens across output chunks.</summary>
    public long OverlapTokenCount { get; set; }

    /// <summary>Total hierarchy-context tokens added to output chunks.</summary>
    public long ContextTokenCount { get; set; }

    /// <summary>Embedding-ready leaf chunks in source order.</summary>
    public IReadOnlyList<ReaderChunk> Chunks { get; set; } = Array.Empty<ReaderChunk>();

    /// <summary>Source span and token evidence aligned to <see cref="Chunks"/>.</summary>
    public IReadOnlyList<ReaderChunkSegment> Segments { get; set; } = Array.Empty<ReaderChunkSegment>();

    /// <summary>Flattened hierarchy nodes with explicit parent/child ids.</summary>
    public IReadOnlyList<ReaderChunkHierarchyNode> Nodes { get; set; } = Array.Empty<ReaderChunkHierarchyNode>();

    /// <summary>Limit and content diagnostics produced during chunking.</summary>
    public IReadOnlyList<OfficeDocumentDiagnostic> Diagnostics { get; set; } = Array.Empty<OfficeDocumentDiagnostic>();
}
