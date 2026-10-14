using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using OfficeIMO.IWork.Internal;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPagesThirdLevelWritingTests {
    [Fact]
    public void Third_decimal_level_keeps_resets_restarts_and_resumed_scope_in_saved_Word() {
        var document = Example(); byte[] bytes = document.SaveBytes();
        Assert.Equal(bytes, document.SaveBytes());
        var source = IWorkSourceDocument.Open(bytes);
        var projection = source.ReadPages();
        Assert.True(projection.Body.IsFormattingComplete);
        Assert.Empty(projection.Diagnostics); Assert.Empty(projection.SourceDeclarationIssues);
        Assert.Empty(projection.SourceReferenceIssues);
        var items = projection.Body.Paragraphs.Where(p => p.ListNumber.HasValue).ToArray();
        Assert.Equal(new[] { 7, 1, 1, 2, 3, 2, 1, 5, 3, 6, 8, 1, 1, 4, 1, 1 }, items.Select(p => p.ListNumber!.Value));
        Assert.Equal(document.Paragraphs.Where(p => p.ListNumber.HasValue).Select(p => p.ListNumber), items.Select(p => p.ListNumber));
        Assert.Equal(new[] { 0, 1, 2, 2, 2, 1, 2, 2, 0, 2, 0, 1, 2, 0, 1, 2 }, items.Select(p => p.ListLevel));
        Assert.All(items, p => Assert.Equal(p.ListLevel * 18, p.ListLayout!.MarkerIndentPoints));
        using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using (var reopened = WordDocument.Load(new MemoryStream(saved.ToArray())))
            Assert.Equal(items.Select(p => p.ListNumber!.Value), WordDocumentTraversal.BuildListIndices(reopened).Values.Select(v => v.Index));
        using var xml = WordprocessingDocument.Open(saved, false);
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(xml));
        var paragraphs = xml.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>()
            .Where(p => p.ParagraphProperties?.NumberingProperties != null).ToArray();
        Assert.Equal(items.Select(p => p.Text), paragraphs.Select(p => p.InnerText));
        for (int i = 0; i < items.Length; i++) {
            var properties = paragraphs[i].ParagraphProperties!;
            Assert.Equal(items[i].ListLevel, properties.NumberingProperties!.NumberingLevelReference!.Val!.Value);
            Assert.Equal(((items[i].ListLevel + 1) * 360).ToString(System.Globalization.CultureInfo.InvariantCulture), properties.Indentation!.Left!.Value);
            Assert.Equal("360", properties.Indentation.Hanging!.Value);
        }
    }

    [Fact]
    public void Third_level_parent_and_overflow_rejections_preserve_list_state() {
        var document = IWorkPagesDocument.Create(); var list = document.CreateNumberedList();
        list.AddItem("Parent", restartAt: 7); byte[] before = document.SaveBytes();
        Assert.Throws<InvalidOperationException>(() => list.AddItem("Missing child", 2));
        Assert.Equal(before, document.SaveBytes());
        list.AddItem("Child", 1); list.AddItem("Grandchild last", 2, int.MaxValue);
        before = document.SaveBytes();
        Assert.Throws<InvalidOperationException>(() => list.AddItem("Grandchild overflow", 2));
        Assert.Equal(before, document.SaveBytes());
        Assert.Equal(2, list.AddItem("Next child", 1).ListNumber);
        Assert.Equal(1, list.AddRichItem(new[] { new IWorkTextRunData("Reset grandchild", "Arial", 12) }, 2).ListNumber);
        Assert.Equal(8, list.AddItem("Next parent").ListNumber);
        before = document.SaveBytes();
        Assert.Throws<InvalidOperationException>(() => list.AddItem("Missing new child", 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => list.AddItem("Outside native levels", 9));
        Assert.Equal(before, document.SaveBytes());
        Assert.Equal(1, list.AddItem("New child", 1).ListNumber);
        Assert.Equal(1, list.AddItem("New grandchild", 2).ListNumber);
    }

    internal static IWorkPagesDocument Example() {
        var document = IWorkPagesDocument.Create();
        document.AddParagraph("Three decimal levels");
        var list = document.CreateNumberedList();
        list.AddItem("Parent seven", restartAt: 7);
        list.AddItem("Child one", 1);
        list.AddItem("Grandchild one", 2); list.AddItem("Grandchild two", 2);
        document.AddParagraph("Plain gap");
        list.AddItem("Grandchild continued three", 2);
        list.AddItem("Child two resets grandchild", 1); list.AddItem("Grandchild reset one", 2);
        list.AddItem("Grandchild restart five", 2, 5);
        document.CreateNumberedList().AddItem("Separate root three", restartAt: 3);
        list.AddItem("Resumed grandchild six", 2);
        list.AddItem("Parent eight resets descendants");
        list.AddItem("Child reset one", 1); list.AddItem("Grandchild reset after root one", 2);
        list.AddItem("Parent restart four", restartAt: 4);
        list.AddItem("Child after restart one", 1); list.AddItem("Grandchild after restart one", 2);
        document.AddParagraph("End third-level profile");
        return document;
    }
}

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Initializing_an_intermediate_counter_clears_an_earlier_explicit_descendant() {
        string[] text = { "Value", "Explicit grandchild", "First child", "Reset grandchild" };
        using var package = ThirdLevelCounterPackage(text, new[] { 0, 2, 1, 2 }, new ulong[] { 7, 5, 0, 0 }, 3);
        var content = ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells.Single().RichText!;
        Assert.True(content.IsFormattingComplete);
        Assert.Equal(new int?[] { 7, 5, 1, 1 }, content.Paragraphs.Select(p => p.ListNumber));
    }

    [Fact]
    public void A_tenth_decimal_level_stays_partial_even_with_complete_decimal_vectors() {
        using var package = ThirdLevelCounterPackage(new[] { "Value" }, new[] { 9 }, new ulong[] { 7 }, 10);
        var content = ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells.Single().RichText!;
        Assert.False(content.IsFormattingComplete);
        Assert.Null(content.Paragraphs.Single().ListNumber);
        Assert.Equal("Value", content.Paragraphs.Single().Text);
    }

    private static MemoryStream ThirdLevelCounterPackage(string[] text, int[] levels, ulong[] starts, int depth) {
        var offsets = ParagraphOffsets(text);
        return SelectedRichTextPackage(IWorkDocumentKind.Numbers,
            Message(StringField(3, text.Length == 1 ? string.Empty : "\n" + string.Join("\n", text.Skip(1))),
                AttributeTable(6, offsets.Select((offset, index) => ListLevelEntry(offset, (ulong)levels[index])).ToArray()),
                AttributeTable(14, offsets.Select((offset, index) => ListLevelEntry(offset, starts[index])).ToArray()),
                AttributeTable(7, AttributeEntry(0, ReferenceField(2, 19)))),
            additionalRecords: ArchiveRecord(19, 2023, Message(Enumerable.Range(0, depth)
                .Select(_ => Message(VarintField(11, 3), VarintField(15, 0), VarintField(25, 0))).ToArray())));
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Third_decimal_level_source_counters_reset_after_each_ancestor_advances(IWorkDocumentKind kind) {
        string[] text = { "Value", "Child one", "Grandchild one", "Grandchild two", "Child two", "Grandchild reset",
            "Parent eight", "Child reset", "Grandchild root reset", "Grandchild restart five", "Grandchild six" };
        int[] level = { 0, 1, 2, 2, 1, 2, 0, 1, 2, 2, 2 };
        var offsets = ParagraphOffsets(text);
        using var package = SelectedRichTextPackage(kind,
            Message(StringField(3, "\n" + string.Join("\n", text.Skip(1))),
                AttributeTable(6, offsets.Select((offset, index) => ListLevelEntry(offset, (ulong)level[index])).ToArray()),
                AttributeTable(14, ListLevelEntry(0, 7), ListLevelEntry(offsets[1], 0), ListLevelEntry(offsets[9], 5),
                    ListLevelEntry(offsets[10], 0)),
                AttributeTable(7, AttributeEntry(0, ReferenceField(2, 19)))),
            additionalRecords: ArchiveRecord(19, 2023, Message(
                VarintField(11, 3), VarintField(11, 3), VarintField(11, 3),
                VarintField(15, 0), VarintField(15, 0), VarintField(15, 0),
                VarintField(25, 0), VarintField(25, 0), VarintField(25, 0))));
        var content = ReadSelectedRichTable(IWorkSourceDocument.Open(package), kind).Item1.Cells.Single().RichText!;
        Assert.True(content.IsFormattingComplete);
        Assert.Equal(new int?[] { 7, 1, 1, 2, 2, 1, 8, 1, 1, 5, 6 }, content.Paragraphs.Select(p => p.ListNumber));
        Assert.Equal(level, content.Paragraphs.Select(p => p.ListLevel));
    }
}
