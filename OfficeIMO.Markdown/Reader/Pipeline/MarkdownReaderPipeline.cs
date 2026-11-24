namespace OfficeIMO.Markdown;

/// <summary>
/// Ordered collection of block parsers that the reader consults at each position.
/// </summary>
internal sealed class MarkdownReaderPipeline {
    private const int DefaultParserCapacity = 20;
    private static readonly IMarkdownBlockParser FrontMatterParser = new MarkdownReader.FrontMatterParser();
    private static readonly IMarkdownBlockParser QuoteParser = new MarkdownReader.QuoteParser();
    private static readonly IMarkdownBlockParser CustomContainerParser = new MarkdownReader.CustomContainerParser();
    private static readonly IMarkdownBlockParser FencedCodeParser = new MarkdownReader.FencedCodeParser();
    private static readonly IMarkdownBlockParser MathBlockParser = new MarkdownReader.MathBlockParser();
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
    private readonly List<MarkdownBlockStart> _blockStarts = new List<MarkdownBlockStart>(DefaultParserCapacity);
    /// <summary>Gets the ordered list of block parsers.</summary>
    public IReadOnlyList<IMarkdownBlockParser> Parsers => _parsers;
    /// <summary>Gets what each parser in <see cref="Parsers"/> needs at the start of a line.</summary>
    public IReadOnlyList<MarkdownBlockStart> BlockStarts => _blockStarts;

    /// <summary>Add a parser to the end of the pipeline.</summary>
    public MarkdownReaderPipeline Add(IMarkdownBlockParser parser) {
        _parsers.Add(parser);
        _blockStarts.Add(MarkdownReader.ClassifyBlockStart(parser));
        return this;
    }
    /// <summary>Insert a parser at the given index in the pipeline.</summary>
    public MarkdownReaderPipeline Insert(int index, IMarkdownBlockParser parser) {
        _parsers.Insert(index, parser);
        _blockStarts.Insert(index, MarkdownReader.ClassifyBlockStart(parser));
        return this;
    }

    /// <summary>Default pipeline covering the syntax OfficeIMO.Markdown emits today.</summary>
    public static MarkdownReaderPipeline Default(MarkdownLoadOptions? options = null) {
        options ??= new MarkdownLoadOptions();
        var sink = new AddingSink(new MarkdownReaderPipeline());
        Compose(options, ref sink);
        return sink.Pipeline;
    }

    // The most recently built pipeline. Pipelines are not changed once built, so parses whose options
    // select the same parsers in the same order share one instead of building their own.
    private static MarkdownReaderPipeline? s_lastDefault;

    /// <summary>Returns the default pipeline for <paramref name="options"/>, reusing the last one built when it matches.</summary>
    internal static MarkdownReaderPipeline GetDefault(MarkdownLoadOptions options) {
        MarkdownReaderPipeline? last = s_lastDefault;
        if (last != null) {
            var matcher = new MatchingSink(last._parsers);
            Compose(options, ref matcher);
            if (matcher.IsMatch) {
                return last;
            }
        }

        MarkdownReaderPipeline pipeline = Default(options);
        s_lastDefault = pipeline;
        return pipeline;
    }

    private interface IParserSink {
        void Add(IMarkdownBlockParser parser);
    }

    private struct AddingSink : IParserSink {
        internal AddingSink(MarkdownReaderPipeline pipeline) {
            Pipeline = pipeline;
        }

        internal MarkdownReaderPipeline Pipeline { get; }

        public void Add(IMarkdownBlockParser parser) => Pipeline.Add(parser);
    }

    private struct MatchingSink : IParserSink {
        private readonly List<IMarkdownBlockParser> _expected;
        private int _count;
        private bool _mismatch;

        internal MatchingSink(List<IMarkdownBlockParser> expected) {
            _expected = expected;
            _count = 0;
            _mismatch = false;
        }

        internal bool IsMatch => !_mismatch && _count == _expected.Count;

        public void Add(IMarkdownBlockParser parser) {
            if (_count >= _expected.Count || !ReferenceEquals(_expected[_count], parser)) {
                _mismatch = true;
            }

            _count++;
        }
    }

    private static void Compose<TSink>(MarkdownLoadOptions options, ref TSink p) where TSink : struct, IParserSink {
        if (options.FrontMatter) p.Add(FrontMatterParser);
        AddExtensions(ref p, options, MarkdownBlockParserPlacement.AfterFrontMatter);
        p.Add(QuoteParser);
        if (options.CustomContainers) p.Add(CustomContainerParser);
        if (options.FencedCode) p.Add(FencedCodeParser);
        if (options.Math) p.Add(MathBlockParser);
        if (options.Images && options.StandaloneImageBlocks) p.Add(ImageParser);
        p.Add(HrParser);
        if (options.HtmlBlocks) p.Add(HtmlBlockParser);
        AddExtensions(ref p, options, MarkdownBlockParserPlacement.AfterHtmlBlocks);
        p.Add(ReferenceLinkDefParser);
        AddExtensions(ref p, options, MarkdownBlockParserPlacement.AfterReferenceLinkDefinitions);
        if (options.Abbreviations) p.Add(AbbreviationDefParser);
        if (options.Tables) p.Add(TableParser);
        if (options.DefinitionLists) p.Add(DefinitionListParser);
        if (options.OrderedLists) p.Add(OrderedListParser);
        if (options.UnorderedLists) p.Add(UnorderedListParser);
        if (options.IndentedCodeBlocks) p.Add(IndentedCodeParser);
        p.Add(SetextHeadingParser);
        if (options.Headings) p.Add(HeadingParser);
        AddExtensions(ref p, options, MarkdownBlockParserPlacement.BeforeParagraphs);
        if (options.Paragraphs) p.Add(ParagraphParser); // must be last
    }

    private static void AddExtensions<TSink>(
        ref TSink pipeline,
        MarkdownLoadOptions options,
        MarkdownBlockParserPlacement placement) where TSink : struct, IParserSink {
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
