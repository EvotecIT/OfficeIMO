namespace OfficeIMO.Markdown;

/// <summary>
/// Ordered collection of block parsers that the reader consults at each position.
/// </summary>
public sealed class MarkdownReaderPipeline {
    private const int DefaultParserCapacity = 20;
    private static readonly IMarkdownBlockParser FrontMatterParser = new MarkdownReader.FrontMatterParser();
    private static readonly IMarkdownBlockParser QuoteParser = new MarkdownReader.QuoteParser();
    private static readonly IMarkdownBlockParser CustomContainerParser = new MarkdownReader.CustomContainerParser();
    private static readonly IMarkdownBlockParser FencedCodeParser = new MarkdownReader.FencedCodeParser();
    private static readonly IMarkdownBlockParser ImageParser = new MarkdownReader.ImageParser();
    private static readonly IMarkdownBlockParser HrParser = new MarkdownReader.HrParser();
    private static readonly IMarkdownBlockParser HtmlBlockParser = new MarkdownReader.HtmlBlockParser();
    private static readonly IMarkdownBlockParser ReferenceLinkDefParser = new MarkdownReader.ReferenceLinkDefParser();
    private static readonly IMarkdownBlockParser AbbreviationDefParser = new MarkdownReader.AbbreviationDefParser();
    private static readonly IMarkdownBlockParser TableParser = new MarkdownReader.TableParser();
    private static readonly IMarkdownBlockParser DefinitionListParser = new MarkdownReader.DefinitionListParser();
    private static readonly IMarkdownBlockParser OrderedListParser = new MarkdownReader.OrderedListParser();
    private static readonly IMarkdownBlockParser UnorderedListParser = new MarkdownReader.UnorderedListParser();
    private static readonly IMarkdownBlockParser IndentedCodeParser = new MarkdownReader.IndentedCodeParser();
    private static readonly IMarkdownBlockParser SetextHeadingParser = new MarkdownReader.SetextHeadingParser();
    private static readonly IMarkdownBlockParser HeadingParser = new MarkdownReader.HeadingParser();
    private static readonly IMarkdownBlockParser ParagraphParser = new MarkdownReader.ParagraphParser();
    private readonly List<IMarkdownBlockParser> _parsers = new List<IMarkdownBlockParser>(DefaultParserCapacity);
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
        if (options.FrontMatter) p.Add(FrontMatterParser);
        AddExtensions(p, options, MarkdownBlockParserPlacement.AfterFrontMatter);
        p.Add(QuoteParser);
        if (options.CustomContainers) p.Add(CustomContainerParser);
        if (options.FencedCode) p.Add(FencedCodeParser);
        if (options.Images && options.StandaloneImageBlocks) p.Add(ImageParser);
        p.Add(HrParser);
        if (options.HtmlBlocks) p.Add(HtmlBlockParser);
        AddExtensions(p, options, MarkdownBlockParserPlacement.AfterHtmlBlocks);
        p.Add(ReferenceLinkDefParser);
        AddExtensions(p, options, MarkdownBlockParserPlacement.AfterReferenceLinkDefinitions);
        if (options.Abbreviations) p.Add(AbbreviationDefParser);
        if (options.Tables) p.Add(TableParser);
        if (options.DefinitionLists) p.Add(DefinitionListParser);
        if (options.OrderedLists) p.Add(OrderedListParser);
        if (options.UnorderedLists) p.Add(UnorderedListParser);
        if (options.IndentedCodeBlocks) p.Add(IndentedCodeParser);
        p.Add(SetextHeadingParser);
        if (options.Headings) p.Add(HeadingParser);
        AddExtensions(p, options, MarkdownBlockParserPlacement.BeforeParagraphs);
        if (options.Paragraphs) p.Add(ParagraphParser); // must be last
        return p;
    }

    private static void AddExtensions(
        MarkdownReaderPipeline pipeline,
        MarkdownReaderOptions options,
        MarkdownBlockParserPlacement placement) {
        var extensions = options.BlockParserExtensions;
        if (extensions.Count == 0) {
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
