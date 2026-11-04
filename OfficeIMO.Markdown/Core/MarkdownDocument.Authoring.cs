namespace OfficeIMO.Markdown;

public partial class MarkdownDocument {
    /// <summary>Appends a paragraph containing plain text and returns it.</summary>
    public ParagraphBlock AddParagraph(string text) => Add(new ParagraphBlock(new InlineSequence().Text(text)));

    /// <summary>Appends a paragraph composed with <see cref="MarkdownParagraphBuilder"/> and returns it.</summary>
    public ParagraphBlock AddParagraph(Action<MarkdownParagraphBuilder> build) {
        if (build == null) throw new ArgumentNullException(nameof(build));
        MarkdownParagraphBuilder builder = new MarkdownParagraphBuilder();
        build(builder);
        return Add(new ParagraphBlock(builder.Inlines));
    }

    /// <summary>Appends an ATX heading and returns it.</summary>
    /// <param name="text">Heading text.</param>
    /// <param name="level">Heading level from 1 to 6.</param>
    public HeadingBlock AddHeading(string text, int level = 1) {
        if (level < 1 || level > 6) throw new ArgumentOutOfRangeException(nameof(level), level, "Markdown heading levels are 1 to 6.");
        return Add(new HeadingBlock(level, text));
    }

    /// <summary>Appends a bulleted list composed with <see cref="MarkdownUnorderedListBuilder"/> and returns it.</summary>
    public UnorderedListBlock AddList(Action<MarkdownUnorderedListBuilder> build) {
        if (build == null) throw new ArgumentNullException(nameof(build));
        MarkdownUnorderedListBuilder builder = new MarkdownUnorderedListBuilder();
        build(builder);
        return Add(builder.Build());
    }

    /// <summary>Appends a bulleted list with one item per value, formatted with <c>ToString()</c>.</summary>
    public UnorderedListBlock AddList<T>(IEnumerable<T> items) {
        if (items == null) throw new ArgumentNullException(nameof(items));
        return AddList(builder => builder.Items(items, null));
    }

    /// <summary>Appends a numbered list composed with <see cref="MarkdownOrderedListBuilder"/> and returns it.</summary>
    public OrderedListBlock AddOrderedList(Action<MarkdownOrderedListBuilder> build) {
        if (build == null) throw new ArgumentNullException(nameof(build));
        MarkdownOrderedListBuilder builder = new MarkdownOrderedListBuilder();
        build(builder);
        return Add(builder.Build());
    }

    /// <summary>Appends a numbered list with one item per value, formatted with <c>ToString()</c>.</summary>
    /// <param name="items">Item values.</param>
    /// <param name="start">First item number.</param>
    public OrderedListBlock AddOrderedList<T>(IEnumerable<T> items, int start = 1) {
        if (items == null) throw new ArgumentNullException(nameof(items));
        return AddOrderedList(builder => builder.StartAt(start).Items(items, null));
    }

    /// <summary>Appends a term/definition list composed with <see cref="MarkdownDefinitionListBuilder"/> and returns it.</summary>
    public DefinitionListBlock AddDefinitionList(Action<MarkdownDefinitionListBuilder> build) {
        if (build == null) throw new ArgumentNullException(nameof(build));
        MarkdownDefinitionListBuilder builder = new MarkdownDefinitionListBuilder();
        build(builder);
        return Add(builder.Build());
    }

    /// <summary>Appends an image and returns it.</summary>
    /// <param name="path">Image path or URL.</param>
    /// <param name="altText">Alternative text.</param>
    /// <param name="title">Optional image title.</param>
    /// <param name="width">Optional display width.</param>
    /// <param name="height">Optional display height.</param>
    public ImageBlock AddImage(string path, string? altText = null, string? title = null, OfficeLength? width = null, OfficeLength? height = null)
        => Add(new ImageBlock(path, altText, title, width, height));

    /// <summary>Appends a table composed with <see cref="MarkdownTableBuilder"/> and returns it.</summary>
    /// <remarks>Column alignment, including the numeric/date heuristics, is set on the builder.</remarks>
    public TableBlock AddTable(Action<MarkdownTableBuilder> build) {
        if (build == null) throw new ArgumentNullException(nameof(build));
        MarkdownTableBuilder builder = new MarkdownTableBuilder();
        build(builder);
        return Add(builder.Build());
    }
}
