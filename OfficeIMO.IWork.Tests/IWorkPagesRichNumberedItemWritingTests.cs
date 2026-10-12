using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPagesRichNumberedItemWritingTests {
    [Fact]
    public void Rich_items_preserve_counters_character_resets_and_plain_markers_in_saved_Word() {
        var document = Example(); byte[] bytes = document.SaveBytes(); Assert.Equal(bytes, document.SaveBytes());
        var source = IWorkSourceDocument.Open(bytes); var projection = source.ReadPages();
        Assert.Empty(projection.Diagnostics); Assert.Empty(projection.SourceDeclarationIssues); Assert.Empty(projection.SourceReferenceIssues);
        var authored = document.Paragraphs.Where(paragraph => paragraph.NumberedList != null).ToArray();
        var items = projection.Body.Paragraphs.Where(paragraph => paragraph.ListNumber.HasValue).ToArray();
        Assert.Equal(new int?[] { 7, 1, 5, 8, 1, 9, 3, 10 }, items.Select(paragraph => paragraph.ListNumber));
        Assert.Equal(authored.Select(paragraph => paragraph.Text), items.Select(paragraph => paragraph.Text));
        for (int index = 0; index < authored.Length; index++) {
            Assert.Equal(authored[index].Runs.Select(run => run.Text), items[index].Runs.Select(run => run.Text));
            Assert.Equal(authored[index].Runs.Select(run => (bool?)run.Bold), items[index].Runs.Select(run => run.Style.Bold));
            Assert.Equal(authored[index].Runs.Select(run => (bool?)run.Italic), items[index].Runs.Select(run => run.Style.Italic));
            Assert.Equal(authored[index].Runs.Select(run => (bool?)run.Underline), items[index].Runs.Select(run => run.Style.Underline));
            Assert.Equal(authored[index].Runs.Select(run => (bool?)run.Strikethrough), items[index].Runs.Select(run => run.Style.Strikethrough));
            Assert.Equal(authored[index].Runs.Select(run => run.Color.RgbHex), items[index].Runs.Select(run => run.Style.Color!.RgbHex));
            var markerDefaults = items[index].Style.TextStyle;
            Assert.Equal(12, markerDefaults.FontSizePoints); Assert.Equal("000000", markerDefaults.Color!.RgbHex);
            Assert.False(markerDefaults.Bold); Assert.False(markerDefaults.Italic);
            Assert.False(markerDefaults.Underline); Assert.False(markerDefaults.Strikethrough);
        }
        using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction(); using var output = new MemoryStream(); result.Value.Save(output);
        using (var reopened = WordDocument.Load(new MemoryStream(output.ToArray())))
            Assert.Equal(items.Select(paragraph => paragraph.ListNumber!.Value),
                WordDocumentTraversal.BuildListIndices(reopened).Values.Select(value => value.Index));
        output.Position = 0; using var saved = WordprocessingDocument.Open(output, false);
        Assert.Empty(new OpenXmlValidator().Validate(saved));
        var numbered = saved.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>()
            .Where(paragraph => paragraph.ParagraphProperties?.NumberingProperties != null).ToArray();
        Assert.Equal(items.Select(paragraph => paragraph.Text), numbered.Select(paragraph => paragraph.InnerText));
        var firstRuns = numbered[0].Elements<W.Run>().Where(run => run.InnerText.Length > 0).ToArray();
        Assert.True(Enabled(firstRuns[0].RunProperties!.Bold)); Assert.False(Enabled(firstRuns[1].RunProperties!.Bold));
        Assert.Equal("003366", firstRuns[0].RunProperties!.Color!.Val!.Value);
        Assert.Equal("000000", firstRuns[1].RunProperties!.Color!.Val!.Value);
        Assert.DoesNotContain(saved.MainDocumentPart.NumberingDefinitionsPart!.Numbering!.Descendants<W.Bold>(), Enabled);
    }

    [Fact]
    public void Rich_runs_are_snapshotted_and_empty_ranges_do_not_change_visible_styling() {
        var document = IWorkPagesDocument.Create(); var list = document.CreateNumberedList();
        var runs = new List<IWorkTextRunData> { new("", fontName: "Times New Roman"), new("Visible", "Arial", 12, bold: true) };
        var paragraph = list.AddRichItem(runs); byte[] baseline = document.SaveBytes(); runs.Clear();
        Assert.Single(paragraph.Runs); Assert.Equal("Visible", paragraph.Text); Assert.Equal(baseline, document.SaveBytes());
        Assert.Throws<NotSupportedException>(() => ((IList<IWorkTextRunData>)paragraph.Runs).Clear());
        var blank = list.AddRichItem(new[] { new IWorkTextRunData("", "Arial", 12, italic: true) });
        Assert.Equal(2, blank.ListNumber); Assert.Single(blank.Runs); Assert.Equal("", blank.Text);
        var projection = IWorkSourceDocument.Open(document.SaveBytes()).ReadPages();
        Assert.Equal(new int?[] { 1, 2 }, projection.Body.Paragraphs.Where(item => item.ListNumber.HasValue).Select(item => item.ListNumber));
    }

    [Fact]
    public void Invalid_rich_runs_preserve_document_and_counter_state() {
        var document = IWorkPagesDocument.Create(); var list = document.CreateNumberedList();
        var valid = new[] { new IWorkTextRunData("Item", "Arial", 12) };
        Assert.Throws<InvalidOperationException>(() => list.AddRichItem(valid, level: 1));
        list.AddItem("Before", restartAt: 7); byte[] baseline = document.SaveBytes();
        Assert.Throws<ArgumentNullException>(() => list.AddRichItem(null!));
        Assert.Throws<ArgumentException>(() => list.AddRichItem(Array.Empty<IWorkTextRunData>()));
        Assert.Throws<ArgumentException>(() => list.AddRichItem(new IWorkTextRunData[] { null! }));
        foreach (string text in new[] { "Line\nbreak", "Line\u2028break", "Line\u2029break" })
            Assert.Throws<ArgumentException>(() => list.AddRichItem(new[] { new IWorkTextRunData(text, "Arial", 12) }, restartAt: 100));
        Assert.Throws<ArgumentException>(() => list.AddRichItem(new[] { new IWorkTextRunData("Size", "Arial", 14.5f) }));
        Assert.Throws<ArgumentOutOfRangeException>(() => list.AddRichItem(valid, level: 9));
        Assert.Throws<ArgumentOutOfRangeException>(() => list.AddRichItem(valid, restartAt: 0));
        Assert.Equal(baseline, document.SaveBytes()); Assert.Single(list.Items); Assert.Single(document.BodyBlocks);
        Assert.Equal(8, list.AddRichItem(valid).ListNumber);
    }

    [Fact]
    public void Rich_item_runs_use_the_existing_native_save_budget_before_stream_mutation() {
        var document = IWorkPagesDocument.Create(); var list = document.CreateNumberedList();
        list.AddRichItem(new[] { new IWorkTextRunData("One", "Arial", 12, bold: true),
            new IWorkTextRunData(" two", "Arial", 12), new IWorkTextRunData(" three", "Arial", 12, italic: true) });
        using var output = new MemoryStream(); output.WriteByte(9);
        Assert.Throws<InvalidDataException>(() => document.Save(output, new() { MaximumTextRuns = 2 }));
        Assert.Equal(new byte[] { 9 }, output.ToArray());
        Assert.NotEmpty(document.SaveBytes(new() { MaximumTextRuns = 3 }));
    }

    internal static IWorkPagesDocument Example() {
        var document = IWorkPagesDocument.Create(); document.AddParagraph("Rich numbered items 😀");
        var list = document.CreateNumberedList();
        list.AddRichItem(new[] { new IWorkTextRunData("Bold first 😀", "Arial", 12, "003366", bold: true),
            new IWorkTextRunData(" plain reset", "Arial", 12) }, restartAt: 7);
        list.AddRichItem(new[] { new IWorkTextRunData("Child italic", "Arial", 12, "008000", italic: true),
            new IWorkTextRunData(", underline", "Arial", 12, "003366", underline: true),
            new IWorkTextRunData(", strike", "Arial", 12, strikethrough: true) }, level: 1);
        list.AddRichItem(new[] { new IWorkTextRunData("Combined decoration", "Arial", 12, "660066", bold: true,
            italic: true, underline: true, strikethrough: true), new IWorkTextRunData(" reset", "Arial", 12) }, level: 1, restartAt: 5);
        document.AddParagraph("Plain gap");
        list.AddRichItem(new[] { new IWorkTextRunData("Colored first", "Arial", 12, "CC0000"),
            new IWorkTextRunData(" parent eight", "Arial", 12) });
        list.AddRichItem(new[] { new IWorkTextRunData("Underlined child", "Arial", 12, underline: true),
            new IWorkTextRunData(" reset", "Arial", 12) }, level: 1);
        list.AddItem("Plain parent nine");
        document.CreateNumberedList().AddRichItem(new[] { new IWorkTextRunData("Separate Zażółć café", "Arial", 12, italic: true) }, restartAt: 3);
        list.AddRichItem(new[] { new IWorkTextRunData("Resumed", "Arial", 12, underline: true),
            new IWorkTextRunData(" parent ten", "Arial", 12) });
        return document;
    }

    private static bool Enabled(W.OnOffType? value) => value != null && value.Val?.Value != false;
}
