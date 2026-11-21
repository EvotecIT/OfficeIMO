using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using WordDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordChartEnumerationTests {
    [Fact]
    public void AllChartsListsBodyHeaderFooterAndFootnoteChartsOfALoadedDocument() {
        using var bytes = new MemoryStream();
        using (var document = WordDocument.Create()) {
            AddChart(document.AddParagraph(), "Body");
            AddChart(document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0], "Cell");
            AddChart(document.Sections[0].PageSetup.Header.AddParagraph(), "Header");
            AddChart(document.Sections[0].PageSetup.Footer.AddParagraph(), "Footer");
            WordFootNote note = document.AddParagraph("Revenue").AddFootNote("Source");
            AddChart(note.Paragraphs!.Last(), "Note");
            document.Save(bytes);
        }

        using var loaded = WordDocument.Load(new MemoryStream(bytes.ToArray()));
        Assert.Equal(new[] { "Body", "Cell" }, loaded.Charts.Select(chart => chart.Title));
        Assert.Equal(new[] { "Body", "Cell", "Header", "Footer", "Note" }, loaded.AllCharts.Select(chart => chart.Title));
        Assert.All(loaded.AllCharts, chart =>
            Assert.Equal(OfficeImageExportFormat.Png, chart.ToImage().Format));
    }

    [Fact]
    public void AllChartsIncludesNotesWithAnExplicitNormalType() {
        using var bytes = new MemoryStream();
        using (var document = WordDocument.Create()) {
            AddChart(document.AddParagraph("Revenue").AddFootNote("Source").Paragraphs!.Last(), "Footnote");
            AddChart(document.AddParagraph("Costs").AddEndNote("Source").Paragraphs!.Last(), "Endnote");
            MainDocumentPart main = document._wordprocessingDocument.MainDocumentPart!;
            foreach (FootnoteEndnoteType note in main.FootnotesPart!.Footnotes!.Elements<Footnote>().Cast<FootnoteEndnoteType>()
                         .Concat(main.EndnotesPart!.Endnotes!.Elements<Endnote>())
                         .Where(note => note.Type == null)) {
                note.Type = FootnoteEndnoteValues.Normal;
            }
            document.Save(bytes);
        }

        using var loaded = WordDocument.Load(new MemoryStream(bytes.ToArray()));
        Assert.Equal(new[] { "Footnote", "Endnote" }, loaded.AllCharts.Select(chart => chart.Title));
    }

    [Fact]
    public void ChartCollectionsListEveryChartDrawingInARun() {
        using var bytes = new MemoryStream();
        using (var document = WordDocument.Create()) {
            AddChart(document.AddParagraph(), "Body A");
            AddChart(document.AddParagraph(), "Body B");
            MergeLastTwoChartRuns(document._document.Body!);
            WordHeader header = document.Sections[0].PageSetup.Header!;
            AddChart(header.AddParagraph(), "Header A");
            AddChart(header.AddParagraph(), "Header B");
            MergeLastTwoChartRuns(document._wordprocessingDocument.MainDocumentPart!.HeaderParts.Single().Header!);
            document.Save(bytes);
        }

        using var loaded = WordDocument.Load(new MemoryStream(bytes.ToArray()));
        Assert.Equal(new[] { "Body A", "Body B" }, loaded.Charts.Select(chart => chart.Title));
        Assert.Equal(new[] { "Body A", "Body B" }, loaded.Sections[0].Charts.Select(chart => chart.Title));
        Assert.Equal(new[] { "Body A", "Body B", "Header A", "Header B" }, loaded.AllCharts.Select(chart => chart.Title));
    }

    [Fact]
    public void ChartCollectionsReadOnlyTheSelectedAlternateContentBranch() {
        using var bytes = new MemoryStream();
        using (var document = WordDocument.Create()) {
            AddChart(document.AddParagraph(), "Body choice");
            AddChart(document.AddParagraph(), "Body fallback");
            WrapLastTwoChartParagraphsInAlternateContent(document._document.Body!);
            WordHeader header = document.Sections[0].PageSetup.Header!;
            AddChart(header.AddParagraph(), "Header choice");
            AddChart(header.AddParagraph(), "Header fallback");
            WrapLastTwoChartParagraphsInAlternateContent(document._wordprocessingDocument.MainDocumentPart!.HeaderParts.Single().Header!);
            document.Save(bytes);
        }

        using var loaded = WordDocument.Load(new MemoryStream(bytes.ToArray()));
        Assert.Equal(new[] { "Body fallback" }, loaded.Charts.Select(chart => chart.Title));
        Assert.Equal(new[] { "Body fallback" }, loaded.Sections[0].Charts.Select(chart => chart.Title));
        Assert.Equal(new[] { "Body fallback", "Header fallback" }, loaded.AllCharts.Select(chart => chart.Title));
    }

    [Fact]
    public void ChartCollectionsFindChartsInEveryWrapperAndSkipDeletedContent() {
        using var bytes = new MemoryStream();
        using (var document = WordDocument.Create()) {
            AddWrappedCharts(document._document.Body!, text => document.AddParagraph(text), "Body");
            WordHeader header = document.Sections[0].PageSetup.Header!;
            AddWrappedCharts(document._wordprocessingDocument.MainDocumentPart!.HeaderParts.Single().Header!, text => header.AddParagraph(text), "Header");
            document.Save(bytes);
        }

        string[] body = VisibleWrappers.Select(name => "Body " + name).ToArray();
        string[] headerTitles = VisibleWrappers.Select(name => "Header " + name).ToArray();
        using var loaded = WordDocument.Load(new MemoryStream(bytes.ToArray()));
        Assert.Equal(body, loaded.Charts.Select(chart => chart.Title));
        Assert.Equal(body, loaded.Sections[0].Charts.Select(chart => chart.Title));
        Assert.Equal(body.Concat(headerTitles), loaded.AllCharts.Select(chart => chart.Title));
        Assert.Equal(ParagraphModelWrappers.Select(name => "Body " + name),
            loaded.Paragraphs.Where(paragraph => paragraph.IsChart).Select(paragraph => paragraph.Chart!.Title));
    }

    private static readonly string[] VisibleWrappers = {
        "Run", "RunAlternateContent", "Hyperlink", "SimpleField", "ComplexField", "Inserted", "MovedTo",
        "CustomXml", "ContentControl", "AlternateContent", "BlockContentControl"
    };

    // Wrappers the paragraph model turns into a paragraph of their own. Paragraph-level mc:AlternateContent and block
    // content controls are reachable only through the chart collections.
    private static readonly string[] ParagraphModelWrappers = {
        "Run", "RunAlternateContent", "Hyperlink", "SimpleField", "ComplexField", "Inserted", "MovedTo", "CustomXml", "ContentControl"
    };

    private static void AddWrappedCharts(OpenXmlElement root, Func<string, WordParagraph> addParagraph, string prefix) {
        var wrappers = new (string Name, Func<Run, OpenXmlElement[]> Wrap)[] {
            ("Run", run => new OpenXmlElement[] { run }),
            ("RunAlternateContent", run => {
                WordDrawing drawing = run.Elements<WordDrawing>().Single();
                drawing.Remove();
                run.Append(UnreadableChoiceWithFallback(drawing));
                return new OpenXmlElement[] { run };
            }),
            ("Hyperlink", run => new OpenXmlElement[] { new Hyperlink(run) { Anchor = "top" } }),
            ("SimpleField", run => new OpenXmlElement[] { new SimpleField(run) { Instruction = "QUOTE chart" } }),
            ("ComplexField", run => new OpenXmlElement[] {
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" QUOTE chart ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                run,
                new Run(new FieldChar { FieldCharType = FieldCharValues.End })
            }),
            ("Inserted", run => new OpenXmlElement[] { new InsertedRun(run) { Id = "1", Author = "Tests" } }),
            ("MovedTo", run => new OpenXmlElement[] { new MoveToRun(run) { Id = "2", Author = "Tests" } }),
            ("CustomXml", run => new OpenXmlElement[] { new CustomXmlRun(run) { Element = "item" } }),
            ("ContentControl", run => new OpenXmlElement[] { new SdtRun(new SdtProperties(new SdtId { Val = 1 }), new SdtContentRun(run)) }),
            ("AlternateContent", run => new OpenXmlElement[] { UnreadableChoiceWithFallback(run) }),
            ("BlockContentControl", run => new OpenXmlElement[] { run }),
            ("Deleted", run => new OpenXmlElement[] { new DeletedRun(run) { Id = "3", Author = "Tests" } }),
            ("MovedFrom", run => new OpenXmlElement[] { new MoveFromRun(run) { Id = "4", Author = "Tests" } })
        };

        foreach ((string name, Func<Run, OpenXmlElement[]> wrap) in wrappers) {
            AddChart(addParagraph(""), prefix + " " + name);
            Paragraph paragraph = LastChartParagraph(root);
            Run run = paragraph.Elements<Run>().Single(candidate => candidate.Elements<WordDrawing>().Any());
            OpenXmlElement anchor = new Run();
            run.InsertBeforeSelf(anchor);
            run.Remove();
            foreach (OpenXmlElement element in wrap(run)) anchor.InsertBeforeSelf(element);
            anchor.Remove();
            if (name == "BlockContentControl") {
                var block = new SdtBlock(new SdtProperties(new SdtId { Val = 2 }), new SdtContentBlock());
                paragraph.InsertBeforeSelf(block);
                paragraph.Remove();
                block.SdtContentBlock!.Append(paragraph);
            }
        }
    }

    /// <summary>Wraps content in the fallback of a compatibility block whose only choice OfficeIMO cannot read.</summary>
    private static AlternateContent UnreadableChoiceWithFallback(OpenXmlElement content) {
        var choice = new AlternateContentChoice { Requires = "unknown" };
        choice.AddNamespaceDeclaration("unknown", "urn:officeimo:tests:unknown");
        return new AlternateContent(choice, new AlternateContentFallback(content));
    }

    private static Paragraph LastChartParagraph(OpenXmlElement root) =>
        root.Descendants<Paragraph>().Last(paragraph => paragraph.Descendants<WordDrawing>().Any());

    /// <summary>Puts the second-to-last chart paragraph in a choice OfficeIMO cannot read and the last in the fallback.</summary>
    private static void WrapLastTwoChartParagraphsInAlternateContent(OpenXmlElement root) {
        Paragraph[] paragraphs = root.Descendants<Paragraph>().Where(paragraph => paragraph.Descendants<WordDrawing>().Any()).ToArray();
        Paragraph choiceParagraph = paragraphs[paragraphs.Length - 2], fallbackParagraph = paragraphs[paragraphs.Length - 1];
        var choice = new AlternateContentChoice { Requires = "unknown" };
        choice.AddNamespaceDeclaration("unknown", "urn:officeimo:tests:unknown");
        var alternateContent = new AlternateContent();
        choiceParagraph.InsertBeforeSelf(alternateContent);
        choiceParagraph.Remove();
        fallbackParagraph.Remove();
        choice.Append(choiceParagraph);
        alternateContent.Append(choice, new AlternateContentFallback(fallbackParagraph));
    }

    /// <summary>Moves the last chart drawing into the previous chart's run, leaving one run that holds both.</summary>
    private static void MergeLastTwoChartRuns(OpenXmlElement root) {
        Run[] runs = root.Descendants<Run>().Where(run => run.Elements<WordDrawing>().Any()).ToArray();
        Run first = runs[runs.Length - 2], second = runs[runs.Length - 1];
        WordDrawing drawing = second.Elements<WordDrawing>().Single();
        drawing.Remove();
        first.Append(drawing);
        second.Ancestors<Paragraph>().First().Remove();
    }

    private static void AddChart(WordParagraph paragraph, string title) {
        WordChart chart = paragraph.AddChart(OfficeChartType.ColumnClustered, new OfficeChartData(new[] { "A", "B" },
            new[] { new OfficeSeries("Value", new[] { 1D, 2D }) }), new OfficeChartOptions { Title = title });
        Assert.Equal(title, chart.Title);
    }
}
