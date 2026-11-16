namespace OfficeIMO.Markdown;

/// <summary>
/// Source information of a parsed document that is captured the first time something reads it.
/// </summary>
/// <remarks>
/// <see cref="MarkdownDocument.Parse(string, MarkdownLoadOptions?)"/> builds the semantic document
/// first. Source spans, the syntax tree and <see cref="MarkdownDocument.ParseInfo"/> are captured from
/// a source-backed parse of the same text when one of them is first read, and attached to the objects
/// the caller already holds. The semantic and source-backed parses build the same object structure,
/// so objects pair up by their position in binding order.
/// </remarks>
internal sealed class MarkdownDeferredSourceCapture {
    private readonly MarkdownDocument _document;
    private readonly string _markdown;
    private readonly MarkdownLoadOptions? _options;
    private readonly object _sync = new();
    private MarkdownObjectSequence? _objects;

    private MarkdownDeferredSourceCapture(
        MarkdownDocument document,
        string markdown,
        MarkdownLoadOptions? options,
        MarkdownObjectSequence objects) {
        _document = document;
        _markdown = markdown;
        _options = options;
        _objects = objects;
    }

    /// <summary>
    /// Returns whether a parse with these options builds the same semantic structure with and without
    /// capturing source information. Parser extensions from outside this library are not known to.
    /// Front matter keeps the source positions of its YAML or TOML values, so a document that may
    /// start with it captures source information while parsing.
    /// </summary>
    internal static bool CanDefer(string markdown, MarkdownLoadOptions options) {
        if (options.FrontMatter) {
            int start = markdown.Length > 0 && markdown[0] == (char)0xFEFF ? 1 : 0;
            if (string.CompareOrdinal(markdown, start, FrontMatterBlock.YamlFence, 0, FrontMatterBlock.YamlFence.Length) == 0
                || string.CompareOrdinal(markdown, start, FrontMatterBlock.TomlFence, 0, FrontMatterBlock.TomlFence.Length) == 0) {
                return false;
            }
        }

        if (options.FencedBlockExtensions.Count != 0
            || options.InlineParserExtensions.Count != 0
            || options.InlineTransformExtensions.Count != 0) {
            return false;
        }

        var blockExtensions = options.BlockParserExtensions;
        for (int i = 0; i < blockExtensions.Count; i++) {
            if (blockExtensions[i]?.Parser is not (MarkdownReader.CalloutParser or MarkdownReader.TocParser or MarkdownReader.FootnoteParser)) {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Defers source capture for a semantic document whose objects were recorded in binding order.
    /// Returns false when an object carries generic attributes, which a deferred object does not
    /// report, or keeps source information that is captured only during parsing; the caller then
    /// parses with source capture instead.
    /// </summary>
    internal static bool TryDefer(
        MarkdownDocument document,
        string markdown,
        MarkdownLoadOptions? options,
        MarkdownObjectSequence objects) {
        foreach (MarkdownObject markdownObject in objects) {
            if (markdownObject is FrontMatterBlock || (markdownObject.HasMetadata && !markdownObject.Attributes.IsEmpty)) {
                return false;
            }
        }

        // Tables keep the inline parse of cells read as blocks outside the object tree, and callers
        // reach it through the table, so those objects are recorded after the tree and defer too.
        int boundCount = objects.Count;
        AppendDetachedObjects(objects);
        int index = 0;
        foreach (MarkdownObject markdownObject in objects) {
            if (index++ >= boundCount && markdownObject.HasMetadata && !markdownObject.Attributes.IsEmpty) {
                return false;
            }
        }

        var capture = new MarkdownDeferredSourceCapture(document, markdown, options, objects);
        foreach (MarkdownObject markdownObject in objects) {
            markdownObject.DeferSourceInformation(capture);
        }

        return true;
    }

    /// <summary>
    /// Appends the objects that tables keep outside the bound object tree, with their descendants, in a
    /// fixed order. A just-bound document and a parse of the same text append corresponding objects.
    /// </summary>
    private static void AppendDetachedObjects(MarkdownObjectSequence objects) {
        List<TableBlock>? tables = null;
        foreach (MarkdownObject markdownObject in objects) {
            if (markdownObject is TableBlock table) {
                (tables ??= new List<TableBlock>()).Add(table);
            }
        }

        if (tables == null) {
            return;
        }

        var visited = new HashSet<MarkdownObject>();
        var pending = new Stack<MarkdownObject>();
        foreach (TableBlock table in tables) {
            foreach (MarkdownObject root in table.EnumerateDetachedObjects()) {
                pending.Push(root);
                while (pending.Count > 0) {
                    MarkdownObject markdownObject = pending.Pop();
                    // Bound objects have a parent; detached ones are reached only from their table.
                    if (markdownObject.Parent != null || !visited.Add(markdownObject)) {
                        continue;
                    }

                    objects.Add(markdownObject);
                    var children = MarkdownObjectTreeBinder.GetChildObjects(markdownObject);
                    for (int i = children.Count - 1; i >= 0; i--) {
                        pending.Push(children[i]);
                    }
                }
            }
        }
    }

    /// <summary>Captures the source information once and attaches it to the document's objects.</summary>
    internal void Capture() {
        lock (_sync) {
            MarkdownObjectSequence? recorded = _objects;
            if (recorded == null) {
                return;
            }

            _objects = null;
            MarkdownParseInfo captured = MarkdownReader.ParseSourceBacked(_markdown, _options ?? new MarkdownLoadOptions());
            var capturedObjects = new MarkdownObjectSequence();
            MarkdownObjectTreeBinder.BindDocument(captured.Document, syntaxTree: null, capturedObjects);
            AppendDetachedObjects(capturedObjects);

            // Each captured object points at the object it describes until the capture ends. The
            // document's objects stay deferred until everything is attached, so a reader on another
            // thread waits for the capture instead of seeing part of it.
            int paired = 0;
            bool pairing = true;
            var capturedMetadata = new MetadataBuffer();
            var capturedEnumerator = capturedObjects.GetEnumerator();
            foreach (MarkdownObject markdownObject in recorded) {
                if (pairing && capturedEnumerator.MoveNext() && capturedEnumerator.Current.GetType() == markdownObject.GetType()) {
                    capturedMetadata.Add(capturedEnumerator.Current.ForwardCapturedObject(markdownObject));
                    paired++;
                } else {
                    pairing = false;
                    markdownObject.ClearDeferredSourceInformation();
                }
            }

            var map = new MarkdownCapturedObjectMap();
            map.Retarget(captured.SyntaxTree);
            if (!ReferenceEquals(captured.FinalSyntaxTree, captured.SyntaxTree)) {
                map.Retarget(captured.FinalSyntaxTree);
            }

            var recordedEnumerator = recorded.GetEnumerator();
            capturedEnumerator = capturedObjects.GetEnumerator();
            for (int i = 0; i < paired && recordedEnumerator.MoveNext() && capturedEnumerator.MoveNext(); i++) {
                recordedEnumerator.Current.AdoptCapturedSourceInformation(capturedEnumerator.Current, map);
                MarkdownSourceSideTables.Adopt(recordedEnumerator.Current, capturedEnumerator.Current);
            }

            _document.ClearSemanticAbbreviationDefinitions();
            _ = new MarkdownParseInfo(
                _document,
                captured.SyntaxTree,
                captured.FinalSyntaxTree,
                captured.SourceMarkdown,
                captured.PreservesOriginalMarkdown ? captured.OriginalMarkdown : null,
                captured.PreservesOriginalMarkdown,
                transformDiagnostics: null,
                referenceLinkDefinitions: captured.ReferenceLinkDefinitions,
                abbreviationDefinitions: captured.AbbreviationDefinitions);

            // Publish the syntax nodes last. Syntax children kept by containers can still lead to
            // objects of the captured parse, so those objects get their own metadata back too.
            recordedEnumerator = recorded.GetEnumerator();
            capturedEnumerator = capturedObjects.GetEnumerator();
            var metadataEnumerator = capturedMetadata.GetEnumerator();
            for (int i = 0; i < paired && recordedEnumerator.MoveNext() && capturedEnumerator.MoveNext() && metadataEnumerator.MoveNext(); i++) {
                capturedEnumerator.Current.RestoreCapturedMetadata(metadataEnumerator.Current);
                recordedEnumerator.Current.PublishCapturedMetadata(metadataEnumerator.Current);
            }
        }
    }

    /// <summary>Append-only metadata list stored in chunks below the large object heap threshold.</summary>
    private sealed class MetadataBuffer {
        private const int ChunkSize = 4096;
        private readonly List<object?[]> _chunks = new();
        private int _count;

        internal void Add(object? item) {
            if (_count % ChunkSize == 0) {
                _chunks.Add(new object?[ChunkSize]);
            }

            _chunks[_count / ChunkSize][_count % ChunkSize] = item;
            _count++;
        }

        internal Enumerator GetEnumerator() => new(this);

        internal struct Enumerator {
            private readonly MetadataBuffer _buffer;
            private int _index;

            internal Enumerator(MetadataBuffer buffer) {
                _buffer = buffer;
                _index = -1;
            }

            internal object? Current => _buffer._chunks[_index / ChunkSize][_index % ChunkSize];

            internal bool MoveNext() => ++_index < _buffer._count;
        }
    }
}

/// <summary>
/// Points syntax nodes and objects of a source-backed parse at the objects of the document they describe.
/// A captured object forwards to its counterpart, so pairing needs no lookup table.
/// </summary>
internal sealed class MarkdownCapturedObjectMap {
    private readonly Stack<MarkdownSyntaxNode> _pending = new();

    /// <summary>Points the syntax nodes under <paramref name="root"/> at the corresponding objects.</summary>
    internal void Retarget(MarkdownSyntaxNode? root) {
        if (root == null) {
            return;
        }

        var pending = _pending;
        pending.Push(root);
        while (pending.Count > 0) {
            MarkdownSyntaxNode node = pending.Pop();
            if (node.AssociatedObject is MarkdownObject { CapturedObjectTarget: { } target }) {
                node.RetargetAssociatedObject(target);
            }

            var children = node.Children;
            for (int i = children.Count - 1; i >= 0; i--) {
                pending.Push(children[i]);
            }
        }
    }

    internal IReadOnlyList<MarkdownSyntaxNode>? Retarget(IReadOnlyList<MarkdownSyntaxNode>? nodes) {
        if (nodes != null) {
            for (int i = 0; i < nodes.Count; i++) {
                Retarget(nodes[i]);
            }
        }

        return nodes;
    }
}
