using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPagesNumberedListWritingTests {
    [Fact]
    public void Decimal_groups_levels_restarts_and_continuation_survive_native_readback_and_saved_Word() {
        var document = Example(); byte[] bytes = document.SaveBytes(); Assert.Equal(bytes, document.SaveBytes());
        var source = IWorkSourceDocument.Open(bytes); var projection = source.ReadPages();
        Assert.Empty(projection.Diagnostics); Assert.Empty(projection.SourceDeclarationIssues); Assert.Empty(projection.SourceReferenceIssues);
        Assert.True(projection.Body.IsFormattingComplete);
        IWorkTextParagraph[] items = projection.Body.Paragraphs.Where(paragraph => paragraph.ListNumber.HasValue).ToArray();
        Assert.Equal(document.Paragraphs.Where(paragraph => paragraph.ListNumber.HasValue).Select(paragraph => paragraph.ListNumber),
            items.Select(paragraph => paragraph.ListNumber));
        Assert.Equal(new[] { 7, 8, 9, 1, 2, 10, 1, 5, 3, 6, 11, 1, 4, 5 }, items.Select(paragraph => paragraph.ListNumber!.Value));
        Assert.Equal(new[] { 0, 0, 0, 1, 1, 0, 1, 1, 0, 1, 0, 1, 0, 0 }, items.Select(paragraph => paragraph.ListLevel));
        Assert.All(items, paragraph => {
            Assert.Equal("1.", paragraph.ListLabel); Assert.Equal("Arial", paragraph.ListFontName);
            Assert.Equal(12, paragraph.Style.TextStyle.FontSizePoints);
            Assert.Equal(1, paragraph.ListLayout!.MarkerScale); Assert.Equal(1.5, paragraph.ListLayout.TextIndentEm);
            Assert.Equal(paragraph.ListLevel == 0 ? 0 : 18, paragraph.ListLayout.MarkerIndentPoints);
        });
        Assert.Null(projection.Body.Paragraphs.Last().ListNumber);
        using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using (var reopened = WordDocument.Load(new MemoryStream(output.ToArray())))
            Assert.Equal(items.Select(paragraph => paragraph.ListNumber!.Value),
                WordDocumentTraversal.BuildListIndices(reopened).Values.Select(value => value.Index));
        using var saved = WordprocessingDocument.Open(output, false); Assert.Empty(new OpenXmlValidator().Validate(saved));
        W.Paragraph[] numbered = saved.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>()
            .Where(paragraph => paragraph.ParagraphProperties?.NumberingProperties != null).ToArray();
        Assert.Equal(14, numbered.Length);
        Assert.Equal(items.Select(paragraph => paragraph.Text), numbered.Select(paragraph => paragraph.InnerText));
        for (int index = 0; index < numbered.Length; index++) {
            Assert.Equal(items[index].ListLevel == 0 ? "360" : "720", numbered[index].ParagraphProperties!.Indentation!.Left!.Value);
            Assert.Equal("360", numbered[index].ParagraphProperties!.Indentation!.Hanging!.Value);
        }
        Assert.Single(saved.MainDocumentPart.Document.Body.Elements<W.Table>());
        Assert.Single(saved.MainDocumentPart.Document.Body.Descendants<W.Drawing>());
    }

    [Fact]
    public void Invalid_items_preserve_document_list_and_counters() {
        var document = IWorkPagesDocument.Create(); var list = document.CreateNumberedList();
        Assert.Throws<InvalidOperationException>(() => list.AddItem("Child", 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => list.AddItem("Deep", 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => list.AddItem("Negative", -1));
        Assert.Throws<ArgumentOutOfRangeException>(() => list.AddItem("Zero", restartAt: 0));
        Assert.Throws<ArgumentException>(() => list.AddItem("Line\nbreak"));
        Assert.ThrowsAny<ArgumentException>(() => list.AddItem("Bad\ud800"));
        Assert.Empty(list.Items); Assert.Empty(document.Paragraphs); Assert.Empty(document.BodyBlocks);
        Assert.Equal(1, list.AddItem("First").ListNumber);
        Assert.Throws<ArgumentException>(() => list.AddItem("TAB\twithout layout"));
        Assert.Equal(2, list.AddItem("Second").ListNumber);
        Assert.Equal(1, list.AddItem("Child\tTAB", 1, layout: new()).ListNumber);
        Assert.Throws<NotSupportedException>(() => ((IList<IWorkPagesParagraphData>)list.Items).Clear());
    }

    [Theory]
    [InlineData("First\u2028Second")]
    [InlineData("First\u2029Second")]
    public void Unicode_separators_are_rejected_atomically_by_all_single_paragraph_paths(string text) {
        var document = IWorkPagesDocument.Create(); var list = document.CreateNumberedList();
        list.AddItem("Before", restartAt: 7); byte[] baseline = document.SaveBytes();
        Assert.Throws<ArgumentException>(() => document.AddParagraph(text));
        Assert.Throws<ArgumentException>(() => document.AddParagraph(text, new IWorkPagesParagraphLayout()));
        var runs = new[] { new IWorkTextRunData("Prefix"), new IWorkTextRunData(text) };
        Assert.Throws<ArgumentException>(() => document.AddRichParagraph(runs));
        Assert.Throws<ArgumentException>(() => document.AddRichParagraph(runs, new IWorkPagesParagraphLayout()));
        Assert.Throws<ArgumentException>(() => list.AddItem(text, restartAt: 100));
        Assert.Equal(baseline, document.SaveBytes()); Assert.Single(list.Items);
        Assert.Single(document.Paragraphs); Assert.Single(document.BodyBlocks);
        Assert.Equal(8, list.AddItem("After").ListNumber);
        using var result = IWorkSourceDocument.Open(document.SaveBytes()).ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        Assert.Equal(new int?[] { 7, 8 }, result.Projection.Body.Paragraphs
            .Where(paragraph => paragraph.ListNumber.HasValue).Select(paragraph => paragraph.ListNumber));
    }

    [Fact]
    public void Counter_overflow_requires_restart_and_parent_advancement_resets_children() {
        var document = IWorkPagesDocument.Create(); var list = document.CreateNumberedList();
        Assert.Equal(int.MaxValue, list.AddItem("Last", restartAt: int.MaxValue).ListNumber);
        Assert.Throws<InvalidOperationException>(() => list.AddItem("Overflow")); Assert.Single(list.Items);
        Assert.Equal(int.MaxValue, list.AddItem("Child last", 1, int.MaxValue).ListNumber);
        Assert.Throws<InvalidOperationException>(() => list.AddItem("Child overflow", 1)); Assert.Equal(2, list.Items.Count);
        Assert.Equal(4, list.AddItem("Parent restart", restartAt: 4).ListNumber);
        Assert.Equal(1, list.AddItem("Reset child", 1).ListNumber);
        using var result = IWorkSourceDocument.Open(document.SaveBytes()).ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        Assert.Equal(new int?[] { int.MaxValue, int.MaxValue, 4, 1 }, result.Projection.Body.Paragraphs
            .Where(paragraph => paragraph.ListNumber.HasValue).Select(paragraph => paragraph.ListNumber));
    }

    [Fact]
    public void Used_list_budget_is_atomic_and_unused_lists_do_not_change_native_bytes() {
        var document = IWorkPagesDocument.Create(); document.AddParagraph("Plain"); byte[] baseline = document.SaveBytes();
        for (int index = 0; index < 5; index++) document.CreateNumberedList();
        Assert.Equal(baseline, document.SaveBytes());
        document.NumberedLists[2].AddItem("First"); document.NumberedLists[4].AddItem("Other");
        using var stream = new MemoryStream(); stream.WriteByte(9);
        var options = new IWorkPagesWriteOptions { MaximumNumberedLists = 1 };
        Assert.Throws<InvalidDataException>(() => document.Save(stream, options)); Assert.Equal(new byte[] { 9 }, stream.ToArray());
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SaveBytes(new() { MaximumNumberedLists = 0 }));
        options.MaximumNumberedLists = 2; Assert.NotEmpty(document.SaveBytes(options));
        string path = Path.Combine(Path.GetTempPath(), "iwork-list-bound-" + Guid.NewGuid().ToString("N") + ".pages");
        try {
            File.WriteAllBytes(path, new byte[] { 37 }); options.MaximumNumberedLists = 1;
            Assert.Throws<InvalidDataException>(() => document.Save(path, overwrite: true, options: options)); Assert.Equal(new byte[] { 37 }, File.ReadAllBytes(path));
        } finally { File.Delete(path); }
    }

    [Fact]
    public void Grouping_restarts_and_fonts_affect_identity_and_independent_models_are_deterministic() {
        var first = Example(); var second = Example(); Assert.Equal(first.SaveBytes(), second.SaveBytes());
        var grouped = IWorkPagesDocument.Create(); var list = grouped.CreateNumberedList(); list.AddItem("One"); list.AddItem("Two");
        var split = IWorkPagesDocument.Create(); split.CreateNumberedList().AddItem("One"); split.CreateNumberedList().AddItem("Two", restartAt: 2);
        Assert.NotEqual(grouped.SaveBytes(), split.SaveBytes());
        var restarted = IWorkPagesDocument.Create(); var restart = restarted.CreateNumberedList(); restart.AddItem("One"); restart.AddItem("Two", restartAt: 2);
        Assert.NotEqual(grouped.SaveBytes(), restarted.SaveBytes());
        Assert.Throws<ArgumentOutOfRangeException>(() => first.CreateNumberedList(fontSizePoints: 11.5f));
        Assert.Throws<ArgumentException>(() => first.CreateNumberedList(fontName: " "));
        Assert.Throws<ArgumentOutOfRangeException>(() => first.CreateNumberedList(fontSizePoints: float.NaN));
        Assert.Equal(2, first.NumberedLists.Count);
    }

    internal static IWorkPagesDocument Example() {
        var document = IWorkPagesDocument.Create(); document.AddParagraph("Owned decimal lists 😀");
        var list = document.CreateNumberedList(); list.AddItem("Parent seven", restartAt: 7); list.AddItem("Parent eight");
        document.AddParagraph("Plain gap"); list.AddItem("Continued nine"); list.AddItem("Child one", 1); list.AddItem("Child two", 1);
        document.AddTable(IWorkTableData.Create(1, 1).SetText(1, 1, "Between list items"), "Values", 180, 28);
        list.AddItem("Parent ten"); list.AddItem("Reset child one", 1); list.AddItem("Child restart five", 1, 5);
        var other = document.CreateNumberedList(); other.AddItem("Separate parent three", restartAt: 3);
        document.AddImage(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports",
            "keynote-created-images-alpha.png")), 72, 48, "Between lists");
        list.AddItem("Resumed child six", 1); list.AddItem("Resumed parent eleven"); list.AddItem("Reset child one again", 1);
        list.AddItem("Parent restart four", restartAt: 4); list.AddItem("Parent five", layout: new(1.25f, 3, 6));
        return document;
    }
}
