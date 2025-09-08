using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Root document container and fluent API entrypoint for composing Markdown.
/// </summary>
public class MarkdownDoc {
    private readonly List<IMarkdownBlock> _blocks = new();
    private IMarkdownBlock? _lastBlock;
    private FrontMatterBlock? _frontMatter;

    public static MarkdownDoc Create() => new MarkdownDoc();

    public IReadOnlyList<IMarkdownBlock> Blocks => _blocks;

    public MarkdownDoc Add(IMarkdownBlock block) {
        if (block is FrontMatterBlock fm) {
            _frontMatter = fm;
        } else {
            _blocks.Add(block);
            _lastBlock = block;
        }
        return this;
    }

    public MarkdownDoc FrontMatter(object data) {
        _frontMatter = FrontMatterBlock.FromObject(data);
        return this;
    }

    public MarkdownDoc H1(string text) => Add(new HeadingBlock(1, text));
    public MarkdownDoc H2(string text) => Add(new HeadingBlock(2, text));
    public MarkdownDoc H3(string text) => Add(new HeadingBlock(3, text));
    public MarkdownDoc H4(string text) => Add(new HeadingBlock(4, text));
    public MarkdownDoc H5(string text) => Add(new HeadingBlock(5, text));
    public MarkdownDoc H6(string text) => Add(new HeadingBlock(6, text));

    public MarkdownDoc P(string text) => Add(new ParagraphBlock(new InlineSequence().Text(text)));

    public MarkdownDoc P(Action<ParagraphBuilder> build) {
        ParagraphBuilder builder = new ParagraphBuilder();
        build(builder);
        return Add(new ParagraphBlock(builder.Inlines));
    }

    public MarkdownDoc Callout(string kind, string title, string body) => Add(new CalloutBlock(kind, title, body));

    public MarkdownDoc Ul(Action<UnorderedListBuilder> build) {
        UnorderedListBuilder builder = new UnorderedListBuilder();
        build(builder);
        return Add(builder.Build());
    }

    public MarkdownDoc Code(string language, string content) {
        CodeBlock code = new CodeBlock(language, content);
        Add(code);
        return this;
    }

    public MarkdownDoc Caption(string caption) {
        if (_lastBlock is ICaptionable cap) {
            cap.Caption = caption;
        } else {
            // Fallback: render as simple paragraph text
            P(caption);
        }
        return this;
    }

    public MarkdownDoc Image(string path, string? alt = null, string? title = null) => Add(new ImageBlock(path, alt, title));

    public MarkdownDoc Table(Action<TableBuilder> build) {
        TableBuilder tb = new TableBuilder();
        build(tb);
        return Add(tb.Build());
    }

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

    public string ToHtml() {
        StringBuilder sb = new StringBuilder();
        foreach (IMarkdownBlock block in _blocks) {
            sb.Append(block.RenderHtml());
        }
        return sb.ToString();
    }
}

