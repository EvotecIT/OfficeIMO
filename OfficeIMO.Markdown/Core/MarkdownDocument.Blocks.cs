namespace OfficeIMO.Markdown;

public partial class MarkdownDocument {
    /// <summary>All blocks added to the document (excluding front matter).</summary>
    public IReadOnlyList<IMarkdownBlock> Blocks => _blocks;
    /// <summary>Document-level front matter/header block when present.</summary>
    public FrontMatterBlock? DocumentHeader => _frontMatter as FrontMatterBlock;
    /// <summary>Structured document front matter entries when present.</summary>
    public IReadOnlyList<FrontMatterBlock.Entry> FrontMatterEntries => DocumentHeader?.Entries ?? Array.Empty<FrontMatterBlock.Entry>();
    /// <summary>All top-level document blocks in order, including front matter when present.</summary>
    public IReadOnlyList<IMarkdownBlock> TopLevelBlocks => BuildTopLevelBlocks();
    /// <summary>Whether the document has front matter.</summary>
    public bool HasDocumentHeader => DocumentHeader != null;

    internal MarkdownParseInfo? AttachedParseResult => _parseResult;

    /// <summary>Adds a block instance and returns that same typed block.</summary>
    /// <typeparam name="T">The block type.</typeparam>
    /// <param name="block">Block to append to the document.</param>
    /// <returns>The attached block.</returns>
    public T Add<T>(T block) where T : IMarkdownBlock {
        Append(block);
        return block;
    }

    /// <summary>Adds blocks in source order and binds the resulting object tree once.</summary>
    /// <param name="blocks">Blocks to append. Enumeration and null validation finish before the document changes.</param>
    /// <returns>A read-only snapshot of the added blocks in source order.</returns>
    public IReadOnlyList<IMarkdownBlock> AddRange(IEnumerable<IMarkdownBlock> blocks) {
        if (blocks == null) throw new ArgumentNullException(nameof(blocks));
        IMarkdownBlock[] added = blocks.ToArray();
        if (added.Any(block => block == null)) throw new ArgumentException("Markdown blocks cannot contain null entries.", nameof(blocks));
        foreach (IMarkdownBlock block in added) {
            if (block is IFrontMatterMarkdownBlock frontMatter) {
                _frontMatter = frontMatter;
            } else {
                _blocks.Add(block);
                _lastBlock = block;
            }
        }
        _parseResult = null;
        MarkdownObjectTreeBinder.BindDocument(this);
        return Array.AsReadOnly(added);
    }

    private MarkdownDocument Append(IMarkdownBlock block) {
        if (block == null) throw new ArgumentNullException(nameof(block));
        if (block is IFrontMatterMarkdownBlock frontMatter) {
            _frontMatter = frontMatter;
        } else {
            _blocks.Add(block);
            _lastBlock = block;
        }
        _parseResult = null;
        RequestObjectTreeBinding();
        return this;
    }
}
