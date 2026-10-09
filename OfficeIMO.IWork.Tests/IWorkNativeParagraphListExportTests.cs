using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkEverydayNativeExportTests {
    [Fact]
    public void Native_Pages_spacing_tabs_and_decimal_restarts_survive_strict_DOCX() {
        VerifyManifest("pages-paragraph-list-v15.4.json");
        using var result = WordIWorkConverter.ConvertPagesToWordResult(
            Corpus("native-exports/pages-paragraph-list-v15.4.pages"),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        Assert.True(result.Projection.Body.IsFormattingComplete);
        result.Report.RequireCompleteEditableReconstruction();
        Assert.Equal(new double?[] { 1.5, 2, 1 }, result.Projection.Body.Paragraphs.Take(3)
            .Select(p => p.Style.LineSpacingMultiplier));
        Assert.Equal(new[] { 180d, 216d }, result.Projection.Body.Paragraphs.Take(2)
            .Select(p => Assert.Single(p.Style.TabStops!).PositionPoints));
        Assert.Empty(result.Projection.Body.Paragraphs[2].Style.TabStops!);
        var counters = result.Projection.Body.Paragraphs.Where(p => p.ListLevel >= 0).ToArray();
        Assert.Equal(new int?[] { 7, 8, 3, 4 }, counters.Select(p => p.ListNumber));
        Assert.Equal(new int?[] { 7, null, 3, null }, counters.Select(p => p.ListStartNumber));
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var actual = WordprocessingDocument.Open(saved, false);
        using var native = WordprocessingDocument.Open(Corpus("native-exports/pages-paragraph-list-v15.4.docx"), false);
        var expectedParagraphs = native.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().ToArray();
        var paragraphs = actual.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().ToArray();
        Assert.Equal(expectedParagraphs.Select(Text), paragraphs.Select(Text));
        for (int index = 0; index < 3; index++) {
            Assert.Equal(EffectiveLine(native, expectedParagraphs[index]), EffectiveLine(actual, paragraphs[index]));
            Assert.Equal(EffectiveTabs(native, expectedParagraphs[index]), EffectiveTabs(actual, paragraphs[index]));
        }
        var numbering = actual.MainDocumentPart.NumberingDefinitionsPart!.Numbering!;
        var ids = paragraphs.Skip(3).Take(4).Select(p => p.ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value).ToArray();
        Assert.Equal(ids[0], ids[1]); Assert.Equal(ids[2], ids[3]); Assert.NotEqual(ids[0], ids[2]);
        Assert.Equal(new[] { 7, 3 }, new[] { ids[0], ids[2] }.Select(id => {
            int abstractId = numbering.Elements<W.NumberingInstance>().Single(n => n.NumberID!.Value == id).AbstractNumId!.Val!.Value;
            return numbering.Elements<W.AbstractNum>().Single(n => n.AbstractNumberId!.Value == abstractId)
                .Elements<W.Level>().Single().StartNumberingValue!.Val!.Value;
        }));
        var nativeNumbering = native.MainDocumentPart.NumberingDefinitionsPart!.Numbering!;
        Assert.Equal(new[] { 7, 3 }, new[] { expectedParagraphs[3], expectedParagraphs[5] }.Select(p => {
            int id = p.ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value;
            return nativeNumbering.Elements<W.NumberingInstance>().Single(n => n.NumberID!.Value == id)
                .Elements<W.LevelOverride>().Single().StartOverrideNumberingValue!.Val!.Value;
        }));
        // Apple's selected decimal level uses 393 twips of text offset/hanging.
        var nativeLevel = nativeNumbering.Elements<W.AbstractNum>().SelectMany(n => n.Elements<W.Level>())
            .Single(level => level.LevelIndex!.Value == 0);
        var expectedIndent = nativeLevel.GetFirstChild<W.PreviousParagraphProperties>()!.Indentation!;
        foreach (var paragraph in paragraphs.Skip(3).Take(4)) {
            W.Indentation indent = paragraph.ParagraphProperties!.Indentation!;
            Assert.Equal(expectedIndent.Left!.Value, indent.Left!.Value);
            Assert.Equal(expectedIndent.Hanging!.Value, indent.Hanging!.Value);
            Assert.Null(indent.FirstLine);
        }
        foreach (var level in numbering.Elements<W.AbstractNum>().SelectMany(n => n.Elements<W.Level>())) {
            var indent = level.GetFirstChild<W.PreviousParagraphProperties>()!.Indentation!;
            Assert.Equal(expectedIndent.Left!.Value, indent.Left!.Value);
            Assert.Equal(expectedIndent.Hanging!.Value, indent.Hanging!.Value);
            Assert.Equal(W.LevelSuffixValues.Tab, level.LevelSuffix!.Val!.Value);
        }
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_PAGES_LIST_LAYOUT_APPROXIMATED");
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(actual));
    }

    [Fact]
    public void Reader_blocks_and_markdown_preserve_native_decimal_counters() {
        var source = IWorkSourceDocument.Open(Corpus("native-exports/pages-paragraph-list-v15.4.pages"));
        var output = new OfficeDocumentReadResult { Kind = ReaderInputKind.IWork };
        var reader = new IWorkReadProjection(output, "paragraph-list.pages", new ReaderOptions(),
            new ReaderIWorkOptions(), System.Threading.CancellationToken.None);
        reader.AddPages(source.ReadPages()); reader.Complete(source);
        Assert.Contains("7. List first", output.Markdown);
        Assert.Contains("8. List second", output.Markdown);
        Assert.Contains("3. List restart", output.Markdown);
        Assert.Contains("4. List continuation", output.Markdown);
        var blocks = output.Pages.SelectMany(page => page.Blocks).Where(block => block.Kind == "list-item").ToArray();
        Assert.Equal(new[] { "7.", "8.", "3.", "4." }, blocks.Select(block => block.Marker));
    }

    private static string? EffectiveLine(WordprocessingDocument document, W.Paragraph paragraph) =>
        paragraph.ParagraphProperties?.SpacingBetweenLines?.Line?.Value
        ?? ParagraphBase(document, paragraph)?.SpacingBetweenLines?.Line?.Value;

    private static int[] EffectiveTabs(WordprocessingDocument document, W.Paragraph paragraph) {
        var positions = new SortedSet<int>();
        foreach (var tabs in new[] { ParagraphBase(document, paragraph)?.Tabs, paragraph.ParagraphProperties?.Tabs }) {
            if (tabs == null) continue;
            foreach (var tab in tabs.Elements<W.TabStop>()) {
                int position = tab.Position!.Value;
                if (tab.Val!.Value == W.TabStopValues.Clear) positions.Remove(position);
                else { Assert.Equal(W.TabStopValues.Left, tab.Val.Value); positions.Add(position); }
            }
        }
        return positions.ToArray();
    }

    private static W.StyleParagraphProperties? ParagraphBase(WordprocessingDocument document, W.Paragraph paragraph) {
        string? id = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        return document.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<W.Style>()
            .FirstOrDefault(style => id != null && style.StyleId?.Value == id)?.StyleParagraphProperties;
    }
}
