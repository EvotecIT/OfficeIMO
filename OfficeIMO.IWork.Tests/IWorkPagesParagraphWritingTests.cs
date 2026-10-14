using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.IWork;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPagesParagraphWritingTests {
    [Fact]
    public void Creation_matches_the_independently_decoded_layout_package_edited_and_reopened_by_Pages() {
        Assert.Equal(File.ReadAllBytes(Fixture("pages-created-layout-v15.4.pages")), Example().SaveBytes());
    }

    [Fact]
    public void Native_saved_layout_matches_Apple_Word_text_spacing_and_tabs_and_keeps_rich_resets() {
        string source = Fixture("pages-created-layout-edited-v15.4.pages");
        var projection = IWorkSourceDocument.Open(source).ReadPages();
        Assert.Empty(projection.Diagnostics); Assert.Empty(projection.SourceDeclarationIssues); Assert.Empty(projection.SourceReferenceIssues);
        Assert.Equal("Default reset Native layout edit.", projection.Body.Paragraphs[6].Text);
        Assert.Equal(18, projection.Body.Paragraphs[4].Style.TextStyle.FontSizePoints);
        Assert.True(projection.Body.Paragraphs[5].Runs[0].Style.Bold);
        Assert.False(projection.Body.Paragraphs[5].Runs[1].Style.Bold);
        Assert.Equal(12, projection.Body.Paragraphs[5].Runs[1].Style.FontSizePoints);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(source,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using var saved = WordprocessingDocument.Open(output, false);
        using var native = WordprocessingDocument.Open(Fixture("pages-created-layout-v15.4.docx"), false);
        Assert.Empty(new OpenXmlValidator().Validate(saved));
        var actual = saved.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().ToArray();
        var reference = native.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().ToArray();
        Assert.Equal(reference.Select(Text), actual.Take(7).Select(Text)); Assert.Equal("", Text(actual[7]));
        Assert.Equal(".", saved.MainDocumentPart.DocumentSettingsPart!.Settings!.GetFirstChild<W.DecimalSymbol>()!.Val!.Value);
        Assert.Equal(new[] { "240", "360", "240", "480", "360", "300", "240", "240" },
            actual.Select(paragraph => paragraph.ParagraphProperties!.SpacingBetweenLines!.Line!.Value));
        foreach (int index in new[] { 1, 3, 4, 5 }) {
            var expected = reference[index].ParagraphProperties!.SpacingBetweenLines!;
            var spacing = actual[index].ParagraphProperties!.SpacingBetweenLines!;
            Assert.Equal(expected.Before?.Value ?? "0", spacing.Before!.Value);
            Assert.Equal(expected.After?.Value ?? "0", spacing.After!.Value);
            Assert.Equal(expected.Line!.Value, spacing.Line!.Value);
        }
        Assert.Equal(reference[2].ParagraphProperties!.Tabs!.Elements<W.TabStop>().Select(tab => (tab.Position!.Value, tab.Val!.Value)),
            actual[2].ParagraphProperties!.Tabs!.Elements<W.TabStop>().Select(tab => (tab.Position!.Value, tab.Val!.Value)));
    }

    [Fact]
    public void Spacing_tabs_and_tab_characters_survive_native_readback_and_saved_Word_without_leaking_to_the_next_paragraph() {
        var layout = new IWorkPagesParagraphLayout(1.5f, 6, 9, Stops().AsEnumerable().Reverse());
        var document = IWorkPagesDocument.Create();
        document.AddParagraph("Start\tLeft\tCenter\tRight\t123.45 😀", layout);
        document.AddParagraph("Default reset");
        document.AddRichParagraph(new[] { new IWorkTextRunData("", fontSizePoints: 18) }, new(2, 12, 18));
        byte[] bytes = document.SaveBytes();
        var projection = IWorkSourceDocument.Open(new MemoryStream(bytes)).ReadPages();
        Assert.Empty(projection.Diagnostics); Assert.Empty(projection.SourceDeclarationIssues); Assert.Empty(projection.SourceReferenceIssues);
        var paragraphs = projection.Body.Paragraphs;
        Assert.Equal(document.Paragraphs[0].Text, paragraphs[0].Text);
        Assert.Equal(1.5, paragraphs[0].Style.LineSpacingMultiplier);
        Assert.Equal(6, paragraphs[0].Style.SpaceBeforePoints); Assert.Equal(9, paragraphs[0].Style.SpaceAfterPoints);
        Assert.Equal(new double[] { 36, 72, 108, 144 }, paragraphs[0].Style.TabStops!.Select(stop => stop.PositionPoints));
        Assert.Equal(Stops().Select(stop => stop.Alignment), paragraphs[0].Style.TabStops!.Select(stop => stop.Alignment));
        Assert.Equal(1, paragraphs[1].Style.LineSpacingMultiplier);
        Assert.Equal(0, paragraphs[1].Style.SpaceBeforePoints); Assert.Equal(0, paragraphs[1].Style.SpaceAfterPoints);
        Assert.Empty(paragraphs[1].Style.TabStops!);
        Assert.Equal(2, paragraphs[2].Style.LineSpacingMultiplier); Assert.Equal(12, paragraphs[2].Style.SpaceBeforePoints);
        Assert.Equal(18, paragraphs[2].Style.SpaceAfterPoints); Assert.Equal("", paragraphs[2].Text);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(new MemoryStream(bytes),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using var saved = WordprocessingDocument.Open(output, false);
        Assert.Empty(new OpenXmlValidator().Validate(saved));
        W.Paragraph[] actual = saved.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().ToArray();
        Assert.Equal(4, actual[0].Descendants<W.TabChar>().Count());
        Assert.Equal(new[] { W.TabStopValues.Left, W.TabStopValues.Center, W.TabStopValues.Right, W.TabStopValues.Decimal },
            actual[0].ParagraphProperties!.Tabs!.Elements<W.TabStop>().Select(stop => stop.Val!.Value));
        Assert.Equal(new[] { 720, 1440, 2160, 2880 },
            actual[0].ParagraphProperties!.Tabs!.Elements<W.TabStop>().Select(stop => stop.Position!.Value));
        W.SpacingBetweenLines spacing = actual[0].ParagraphProperties!.SpacingBetweenLines!;
        Assert.Equal("120", spacing.Before!.Value); Assert.Equal("180", spacing.After!.Value);
        Assert.Equal("360", spacing.Line!.Value); Assert.Equal(W.LineSpacingRuleValues.Auto, spacing.LineRule!.Value);
        Assert.Equal("240", actual[1].ParagraphProperties!.SpacingBetweenLines!.Line!.Value);
        Assert.Empty(actual[1].ParagraphProperties!.Tabs?.Elements<W.TabStop>() ?? Enumerable.Empty<W.TabStop>());
    }

    [Fact]
    public void Layout_copies_and_sorts_tabs_and_explicit_defaults_keep_the_existing_native_bytes() {
        var input = Stops().AsEnumerable().Reverse().ToList();
        var layout = new IWorkPagesParagraphLayout(tabStops: input); input.Clear();
        Assert.Equal(new double[] { 36, 72, 108, 144 }, layout.TabStops.Select(stop => stop.PositionPoints));
        Assert.Throws<NotSupportedException>(() => ((IList<IWorkTabStop>)layout.TabStops).Clear());
        var first = IWorkPagesDocument.Create(); first.AddParagraph("Layout", layout);
        var second = IWorkPagesDocument.Create(); second.AddParagraph("Layout", new IWorkPagesParagraphLayout(tabStops: Stops()));
        Assert.Equal(first.SaveBytes(), second.SaveBytes());
        var original = IWorkPagesDocument.Create(); original.AddParagraph("Default");
        var explicitDefault = IWorkPagesDocument.Create(); explicitDefault.AddParagraph("Default", new IWorkPagesParagraphLayout());
        Assert.Equal(original.SaveBytes(), explicitDefault.SaveBytes());
    }

    [Fact]
    public void Aggregate_custom_tab_budget_fails_before_destination_publication_and_accepts_the_exact_boundary() {
        var document = IWorkPagesDocument.Create();
        var layout = new IWorkPagesParagraphLayout(tabStops: Stops());
        document.AddParagraph("First\tvalue", layout); document.AddParagraph("Second\tvalue", layout);
        var limits = new IWorkPagesWriteOptions { MaximumTabStops = 7 };
        using var output = new MemoryStream(new byte[] { 9, 8, 7 });
        Assert.Throws<InvalidDataException>(() => document.Save(output, limits)); Assert.Equal(new byte[] { 9, 8, 7 }, output.ToArray());
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-IWork-Paragraph-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try {
            string path = Path.Combine(folder, "existing.pages"); File.WriteAllBytes(path, new byte[] { 9, 8, 7 });
            Assert.Throws<InvalidDataException>(() => document.Save(path, overwrite: true, options: limits));
            Assert.Equal(new byte[] { 9, 8, 7 }, File.ReadAllBytes(path)); Assert.Single(Directory.GetFiles(folder));
            limits.MaximumTabStops = 8;
            var projection = IWorkSourceDocument.Open(new MemoryStream(document.SaveBytes(limits))).ReadPages();
            Assert.All(projection.Body.Paragraphs.Take(2), paragraph => Assert.Equal(4, paragraph.Style.TabStops!.Count));
            using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
            Assert.Throws<OperationCanceledException>(() => document.Save(path, overwrite: true, options: limits, cancellationToken: cancelled.Token));
            Assert.Equal(new byte[] { 9, 8, 7 }, File.ReadAllBytes(path));
        } finally { Directory.Delete(folder, recursive: true); }
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SaveBytes(new() { MaximumTabStops = 0 }));
    }

    [Fact]
    public void Invalid_layouts_and_paragraphs_fail_without_model_mutation_and_Keynote_rejects_reused_tabbed_runs() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new IWorkPagesParagraphLayout(0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new IWorkPagesParagraphLayout(float.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => new IWorkPagesParagraphLayout(spaceBeforePoints: -1));
        Assert.Throws<ArgumentOutOfRangeException>(() => new IWorkPagesParagraphLayout(spaceAfterPoints: float.PositiveInfinity));
        Assert.Throws<ArgumentOutOfRangeException>(() => IWorkTabStop.Create(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => IWorkTabStop.Create(float.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => IWorkTabStop.Create(1, (IWorkTabAlignment)(-1)));
        Assert.Throws<ArgumentException>(() => new IWorkPagesParagraphLayout(tabStops: new[] { IWorkTabStop.Create(1), IWorkTabStop.Create(1, IWorkTabAlignment.Right) }));
        Assert.Throws<ArgumentException>(() => new IWorkPagesParagraphLayout(tabStops: new IWorkTabStop[] { null! }));
        var document = IWorkPagesDocument.Create();
        var layout = new IWorkPagesParagraphLayout();
        Assert.Throws<ArgumentNullException>(() => document.AddParagraph("Value", layout: null!));
        Assert.Throws<ArgumentException>(() => document.AddParagraph("one\r\ntwo", layout));
        Assert.Throws<ArgumentException>(() => document.AddParagraph("one\u0001two", layout));
        Assert.Empty(document.Paragraphs);
        var paragraph = document.AddParagraph("Value\there", layout);
        var keynote = IWorkKeynoteDocument.Create(); var slide = keynote.AddSlide();
        Assert.Throws<ArgumentException>(() => slide.AddRichText(paragraph.Runs, 0, 0, 200, 100));
        Assert.Empty(slide.TextBoxes);
    }

    private static IWorkTabStop[] Stops() => Enumerable.Range(0, 4)
        .Select(index => IWorkTabStop.Create((index + 1) * 36, (IWorkTabAlignment)index)).ToArray();

    internal static IWorkPagesDocument Example() {
        var document = IWorkPagesDocument.Create();
        document.AddParagraph("Native paragraph layout", fontSizePoints: 18);
        document.AddParagraph("Relative spacing 1.5: this paragraph has six points before and nine points after. Its text wraps across lines so the native line spacing can be inspected alongside the exported paragraph settings.",
            new IWorkPagesParagraphLayout(1.5f, 6, 9));
        document.AddParagraph("Start\tLeft\tCenter\tRight\t123.45", new IWorkPagesParagraphLayout(tabStops: new[] {
            IWorkTabStop.Create(90), IWorkTabStop.Create(180, IWorkTabAlignment.Center),
            IWorkTabStop.Create(270, IWorkTabAlignment.Right), IWorkTabStop.Create(360, IWorkTabAlignment.Decimal)
        }));
        document.AddParagraph("Double spacing: this paragraph has twelve points before and eighteen points after. Wrapped lines make the relative spacing visible without changing the font or body width.",
            new IWorkPagesParagraphLayout(2, 12, 18));
        document.AddParagraph("", new IWorkPagesParagraphLayout(1.5f, 6, 6), fontSizePoints: 18);
        document.AddRichParagraph(new[] {
            new IWorkTextRunData("Rich layout", fontSizePoints: 14, bold: true),
            new IWorkTextRunData(" plain Zażółć 😀", fontSizePoints: 12)
        }, new IWorkPagesParagraphLayout(1.25f, spaceAfterPoints: 6));
        document.AddParagraph("Default reset paragraph.");
        return document;
    }

    private static string Text(W.Paragraph paragraph) => string.Concat(paragraph.Descendants<W.Run>()
        .SelectMany(run => run.ChildElements).Select(element => element is W.Text text ? text.Text : element is W.TabChar ? "\t" : ""));

    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports", name);
}
