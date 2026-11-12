namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPagesHalfPointListTests {
    [Theory]
    [InlineData(11.5f)]
    [InlineData(12.5f)]
    public void Uniform_half_point_items_keep_marker_and_body_precision_in_saved_Word(float size) {
        var document = IWorkPagesDocument.Create();
        var list = document.CreateNumberedList(fontSizePoints: size);
        list.AddItem("Parent", restartAt: 7); list.AddItem("Child", level: 1);
        var source = IWorkSourceDocument.Load(document.ToBytes());
        Assert.Empty(source.ReadPages().SourceDeclarationIssues);
        IWorkPagesMixedSizeListTests.CheckWord(source);
    }

    [Fact]
    public void Mixed_half_point_runs_keep_first_run_markers_nine_levels_and_resumed_counters() {
        var document = Example(); byte[] bytes = document.ToBytes();
        Assert.Equal(bytes, document.ToBytes());
        var source = IWorkSourceDocument.Load(bytes); var projection = source.ReadPages();
        Assert.True(projection.Body.IsFormattingComplete);
        Assert.Empty(projection.Diagnostics); Assert.Empty(projection.SourceDeclarationIssues);
        Assert.Empty(projection.SourceReferenceIssues);
        var authored = document.Paragraphs.Where(item => item.NumberedList != null).ToArray();
        var items = projection.Body.Paragraphs.Where(item => item.ListNumber.HasValue).ToArray();
        Assert.Equal(authored.Select(item => item.ListNumber), items.Select(item => item.ListNumber));
        Assert.Equal(authored.Select(item => item.ListLevel), items.Select(item => item.ListLevel));
        for (int index = 0; index < items.Length; index++) {
            Assert.Equal(11.5, items[index].Style.TextStyle.FontSizePoints);
            Assert.Equal(authored[index].Runs.Select(run => (double)run.FontSizePoints),
                items[index].Runs.Select(run => run.Style.FontSizePoints!.Value));
        }
        IWorkPagesMixedSizeListTests.CheckWord(source);
    }

    [Fact]
    public void Finer_list_sizes_reject_before_document_or_counter_mutation() {
        var document = IWorkPagesDocument.Create();
        var list = document.CreateNumberedList(fontSizePoints: 11.5f);
        list.AddItem("Before", restartAt: 7); byte[] original = document.ToBytes();
        Assert.Throws<ArgumentOutOfRangeException>(() => document.CreateNumberedList(fontSizePoints: 11.25f));
        Assert.Single(document.NumberedLists);
        Assert.Throws<ArgumentException>(() => list.AddRichItem(new[] { new IWorkTextRunData("Quarter", fontSizePoints: 14.25f) }, restartAt: 100));
        Assert.Equal(original, document.ToBytes()); Assert.Single(list.Items);
        Assert.Equal(8, list.AddItem("After").ListNumber);
    }

    internal static IWorkPagesDocument Example() {
        var document = IWorkPagesDocument.Create(); document.AddParagraph("Half point decimal items");
        var primary = document.CreateNumberedList("Arial", 11.5f);
        primary.AddRichItem(Runs("Root seven", 18.5f, 8.5f, 14.5f), restartAt: 7);
        for (int level = 1; level <= 8; level++) primary.AddRichItem(Runs("Depth " + level), level: level);
        primary.AddRichItem(Runs("Deep two"), level: 8);
        document.AddParagraph("Plain gap retains counters");
        primary.AddRichItem(Runs("Root eight", 8.5f, 18.5f, 14.5f));
        primary.AddRichItem(Runs("Child one", 8.5f, 10.5f, 14.5f), level: 1);
        var serif = document.CreateNumberedList(IWorkDecimalListMarkerStyle.Parentheses, "Times New Roman", 11.5f);
        serif.AddRichItem(Runs("Serif marker"), restartAt: 3); serif.AddRichItem(Runs("Serif child"), level: 1);
        var mono = document.CreateNumberedList(IWorkDecimalListMarkerStyle.ClosingParenthesis, "Courier New", 11.5f);
        mono.AddRichItem(Runs("Mono marker"), restartAt: 5); mono.AddRichItem(Runs("Mono child"), level: 1);
        primary.AddRichItem(Runs("Resumed child", 12.5f, 24.5f, 8.5f), level: 1);
        var uniform = document.CreateNumberedList("Arial", 11.5f);
        uniform.AddRichItem(Runs("Uniform root", 11.5f, 11.5f, 11.5f), restartAt: 11);
        uniform.AddRichItem(Runs("Uniform child", 11.5f, 11.5f, 11.5f), level: 1);
        return document;
    }

    private static IWorkTextRunData[] Runs(string text, float first = 12.5f, float middle = 12.5f, float last = 12.5f) => new[] {
        new IWorkTextRunData(text + " serif", "Times New Roman", first, "003366", bold: true),
        new IWorkTextRunData(" mono", "Courier New", middle, "008000", italic: true),
        new IWorkTextRunData(" plain", "Arial", last)
    };
}
