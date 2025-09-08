using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Root document container and fluent API entrypoint for composing Markdown.
/// Supports a fluent-chaining style and an explicit object model via <see cref="Add(IMarkdownBlock)"/>.
/// </summary>
public class MarkdownDoc {
    private readonly List<IMarkdownBlock> _blocks = new();
    private IMarkdownBlock? _lastBlock;
    private FrontMatterBlock? _frontMatter;

    /// <summary>Creates a new, empty Markdown document.</summary>
    public static MarkdownDoc Create() => new MarkdownDoc();

    /// <summary>All blocks added to the document (excluding front matter).</summary>
    public IReadOnlyList<IMarkdownBlock> Blocks => _blocks;

    /// <summary>Adds a block instance (object-model style).</summary>
    /// <param name="block">Block to append to the document.</param>
    /// <returns>Same <see cref="MarkdownDoc"/> for chaining.</returns>
    public MarkdownDoc Add(IMarkdownBlock block) {
        if (block is FrontMatterBlock fm) {
            _frontMatter = fm;
        } else {
            _blocks.Add(block);
            _lastBlock = block;
        }
        return this;
    }

    /// <summary>Sets YAML front matter from an anonymous object or dictionary.</summary>
    public MarkdownDoc FrontMatter(object data) {
        _frontMatter = FrontMatterBlock.FromObject(data);
        return this;
    }

    /// <summary>Adds an H1 heading.</summary>
    public MarkdownDoc H1(string text) => Add(new HeadingBlock(1, text));
    /// <summary>Adds an H2 heading.</summary>
    public MarkdownDoc H2(string text) => Add(new HeadingBlock(2, text));
    /// <summary>Adds an H3 heading.</summary>
    public MarkdownDoc H3(string text) => Add(new HeadingBlock(3, text));
    /// <summary>Adds an H4 heading.</summary>
    public MarkdownDoc H4(string text) => Add(new HeadingBlock(4, text));
    /// <summary>Adds an H5 heading.</summary>
    public MarkdownDoc H5(string text) => Add(new HeadingBlock(5, text));
    /// <summary>Adds an H6 heading.</summary>
    public MarkdownDoc H6(string text) => Add(new HeadingBlock(6, text));

    /// <summary>Adds a paragraph with plain text.</summary>
    public MarkdownDoc P(string text) => Add(new ParagraphBlock(new InlineSequence().Text(text)));

    /// <summary>Adds a paragraph composed with the paragraph builder.</summary>
    public MarkdownDoc P(Action<ParagraphBuilder> build) {
        ParagraphBuilder builder = new ParagraphBuilder();
        build(builder);
        return Add(new ParagraphBlock(builder.Inlines));
    }

    /// <summary>Adds a callout/admonition block (Docs-style).</summary>
    public MarkdownDoc Callout(string kind, string title, string body) => Add(new CalloutBlock(kind, title, body));

    /// <summary>Adds an unordered list.</summary>
    public MarkdownDoc Ul(Action<UnorderedListBuilder> build) {
        UnorderedListBuilder builder = new UnorderedListBuilder();
        build(builder);
        return Add(builder.Build());
    }

    /// <summary>Adds an ordered list.</summary>
    public MarkdownDoc Ol(Action<OrderedListBuilder> build) {
        OrderedListBuilder builder = new OrderedListBuilder();
        build(builder);
        return Add(builder.Build());
    }

    /// <summary>Adds a definition list.</summary>
    public MarkdownDoc Dl(Action<DefinitionListBuilder> build) {
        DefinitionListBuilder builder = new DefinitionListBuilder();
        build(builder);
        return Add(builder.Build());
    }

    /// <summary>Adds a fenced code block.</summary>
    public MarkdownDoc Code(string language, string content) {
        CodeBlock code = new CodeBlock(language, content);
        Add(code);
        return this;
    }

    /// <summary>Sets a caption for the last captionable block (image/code), or appends a paragraph.</summary>
    public MarkdownDoc Caption(string caption) {
        if (_lastBlock is ICaptionable cap) {
            cap.Caption = caption;
        } else {
            // Fallback: render as simple paragraph text
            P(caption);
        }
        return this;
    }

    /// <summary>Adds an image block with optional alt text and title.</summary>
    public MarkdownDoc Image(string path, string? alt = null, string? title = null) => Add(new ImageBlock(path, alt, title));

    /// <summary>Adds a table built with <see cref="TableBuilder"/>.</summary>
    public MarkdownDoc Table(Action<TableBuilder> build) {
        TableBuilder tb = new TableBuilder();
        build(tb);
        return Add(tb.Build());
    }

    /// <summary>Renders the document to Markdown string.</summary>
    public string ToMarkdown() {
        StringBuilder sb = new StringBuilder();
        if (_frontMatter != null) {
            sb.AppendLine(_frontMatter.Render());
            sb.AppendLine();
        }
        for (int i = 0; i < _blocks.Count; i++) {
            string rendered = _blocks[i].RenderMarkdown();
            if (!string.IsNullOrEmpty(rendered)) sb.AppendLine(rendered);
            if (i < _blocks.Count - 1) sb.AppendLine();
        }
        return sb.ToString();
    }

    /// <summary>Renders a basic HTML representation of the document (no front matter).</summary>
    public string ToHtml() {
        StringBuilder sb = new StringBuilder();
        foreach (IMarkdownBlock block in _blocks) {
            sb.Append(block.RenderHtml());
        }
        return sb.ToString();
    }
}
