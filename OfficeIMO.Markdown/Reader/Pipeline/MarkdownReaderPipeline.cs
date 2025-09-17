using System.Collections.Generic;

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
    public static MarkdownReaderPipeline Default() {
        var p = new MarkdownReaderPipeline();
        p.Add(new MarkdownReader.FrontMatterParser());
        p.Add(new MarkdownReader.CalloutParser());
        p.Add(new MarkdownReader.QuoteParser());
        p.Add(new MarkdownReader.FencedCodeParser());
        p.Add(new MarkdownReader.ImageParser());
        p.Add(new MarkdownReader.HrParser());
        p.Add(new MarkdownReader.HtmlBlockParser());
        p.Add(new MarkdownReader.TocParser());
        p.Add(new MarkdownReader.ReferenceLinkDefParser());
        p.Add(new MarkdownReader.FootnoteParser());
        p.Add(new MarkdownReader.TableParser());
        p.Add(new MarkdownReader.DefinitionListParser());
        p.Add(new MarkdownReader.OrderedListParser());
        p.Add(new MarkdownReader.UnorderedListParser());
        p.Add(new MarkdownReader.SetextHeadingParser());
        p.Add(new MarkdownReader.HeadingParser());
        p.Add(new MarkdownReader.ParagraphParser()); // must be last
        return p;
    }
}
