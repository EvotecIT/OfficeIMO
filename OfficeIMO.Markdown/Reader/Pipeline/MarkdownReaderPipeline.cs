namespace OfficeIMO.Markdown;

/// <summary>
/// Ordered collection of block parsers that the reader consults at each position.
/// </summary>
public sealed class MarkdownReaderPipeline {
    private readonly List<IMarkdownBlockParser> _parsers = new List<IMarkdownBlockParser>();
    /// <summary>Gets the ordered list of block parsers.</summary>
    public IReadOnlyList<IMarkdownBlockParser> Parsers => _parsers;

    /// <summary>Add a parser to the end of the pipeline.</summary>
    public MarkdownReaderPipeline Add(IMarkdownBlockParser parser) { _parsers.Add(parser); return this; }
    /// <summary>Insert a parser at the given index in the pipeline.</summary>
    public MarkdownReaderPipeline Insert(int index, IMarkdownBlockParser parser) { _parsers.Insert(index, parser); return this; }

    /// <summary>Default pipeline covering the syntax OfficeIMO.Markdown emits today.</summary>
    public static MarkdownReaderPipeline Default(MarkdownReaderOptions? options = null) {
        options ??= new MarkdownReaderOptions();
        var p = new MarkdownReaderPipeline();
        if (options.FrontMatter) p.Add(new MarkdownReader.FrontMatterParser());
        AddExtensions(p, options, MarkdownBlockParserPlacement.AfterFrontMatter);
        p.Add(new MarkdownReader.QuoteParser());
        if (options.FencedCode) p.Add(new MarkdownReader.FencedCodeParser());
        if (options.Images && options.StandaloneImageBlocks) p.Add(new MarkdownReader.ImageParser());
        p.Add(new MarkdownReader.HrParser());
        if (options.HtmlBlocks) p.Add(new MarkdownReader.HtmlBlockParser());
        AddExtensions(p, options, MarkdownBlockParserPlacement.AfterHtmlBlocks);
        p.Add(new MarkdownReader.ReferenceLinkDefParser());
        AddExtensions(p, options, MarkdownBlockParserPlacement.AfterReferenceLinkDefinitions);
        if (options.Tables) p.Add(new MarkdownReader.TableParser());
        if (options.DefinitionLists) p.Add(new MarkdownReader.DefinitionListParser());
        if (options.OrderedLists) p.Add(new MarkdownReader.OrderedListParser());
        if (options.UnorderedLists) p.Add(new MarkdownReader.UnorderedListParser());
        if (options.IndentedCodeBlocks) p.Add(new MarkdownReader.IndentedCodeParser());
        p.Add(new MarkdownReader.SetextHeadingParser());
        if (options.Headings) p.Add(new MarkdownReader.HeadingParser());
        AddExtensions(p, options, MarkdownBlockParserPlacement.BeforeParagraphs);
        if (options.Paragraphs) p.Add(new MarkdownReader.ParagraphParser()); // must be last
        return p;
    }

    private static void AddExtensions(
        MarkdownReaderPipeline pipeline,
        MarkdownReaderOptions options,
        MarkdownBlockParserPlacement placement) {
        var extensions = options.BlockParserExtensions;
        if (extensions == null || extensions.Count == 0) {
            return;
        }

        for (int i = 0; i < extensions.Count; i++) {
            var extension = extensions[i];
            if (extension == null
                || extension.Placement != placement
                || !extension.AppliesTo(options)) {
                continue;
            }

            pipeline.Add(extension.Parser);
        }
    }
}
