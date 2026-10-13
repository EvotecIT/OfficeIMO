using System.Text;
using System.Threading;
using OfficeIMO.Reader.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPagesNineLevelWritingTests {
    internal static readonly string[] LevelNames = { "Root seven", "Level one", "Level two", "Level three",
        "Level four", "Level five", "Level six", "Level seven", "Level eight" };
    internal static readonly int[] Numbers = { 7, 1, 1, 1, 1, 1, 1, 1, 1, 2, 3, 4, 1, 1, 1, 1, 9, 3, 10,
        8, 1, 1, 1, 1, 1, 1, 1, 1, 4, 1, 1, 1, 1, 1, 1, 1, 1 };
    internal static readonly int[] Levels = { 0, 1, 2, 3, 4, 5, 6, 7, 8, 8, 8, 4, 5, 6, 7, 8, 8, 0, 8,
        0, 1, 2, 3, 4, 5, 6, 7, 8, 0, 1, 2, 3, 4, 5, 6, 7, 8 };

    [Fact]
    public void Nine_native_decimal_levels_retain_resets_restarts_and_resumed_scope_in_saved_Word() {
        var model = Example(); byte[] bytes = model.SaveBytes();
        Assert.Equal(bytes, model.SaveBytes());
        var source = IWorkSourceDocument.Open(bytes); var projection = source.ReadPages();
        Assert.True(projection.Body.IsFormattingComplete); Assert.Empty(projection.Diagnostics);
        Assert.Empty(projection.SourceDeclarationIssues); Assert.Empty(projection.SourceReferenceIssues);
        var items = projection.Body.Paragraphs.Where(p => p.ListNumber.HasValue).ToArray();
        Assert.Equal(Numbers, items.Select(p => p.ListNumber!.Value));
        Assert.Equal(Levels, items.Select(p => p.ListLevel));
        Assert.Equal(model.Paragraphs.Select(p => p.Text), projection.Body.Paragraphs.Take(model.Paragraphs.Count).Select(p => p.Text));
        Assert.All(items, p => {
            Assert.Equal(p.ListLevel * 18, p.ListLayout!.MarkerIndentPoints);
            Assert.Equal(new string(' ', p.ListLevel * 2) + p.ListNumber + ". " + p.Text,
                IWorkReadProjection.RichTextMarkdown(p));
        });
        IWorkPagesDecimalTextEditingTests.CheckWord(source);
    }

    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(2)] [InlineData(3)]
    [InlineData(4)] [InlineData(5)] [InlineData(6)] [InlineData(7)]
    public void Advancing_each_ancestor_resets_all_its_descendants_and_requires_new_immediate_parents(int ancestor) {
        var model = IWorkPagesDocument.Create(); var list = model.CreateNumberedList();
        for (int level = 0; level < 9; level++) list.AddItem("Initial " + level, level, 5);
        Assert.Equal(6, list.AddItem("Advance ancestor", ancestor).ListNumber);
        byte[] before = model.SaveBytes();
        if (ancestor < 7)
            Assert.Throws<InvalidOperationException>(() => list.AddItem("Missing immediate parent", ancestor + 2));
        Assert.Equal(before, model.SaveBytes());
        for (int level = ancestor + 1; level < 9; level++)
            Assert.Equal(1, list.AddItem("Reset descendant " + level, level).ListNumber);
        var projection = IWorkSourceDocument.Open(model.SaveBytes()).ReadPages();
        Assert.True(projection.Body.IsFormattingComplete);
        Assert.Equal(list.Items.Select(p => p.ListNumber), projection.Body.Paragraphs.Take(list.Items.Count).Select(p => p.ListNumber));
        IWorkPagesDecimalTextEditingTests.CheckWord(IWorkSourceDocument.Open(model.SaveBytes()));
    }

    [Fact]
    public void Deepest_counter_overflow_and_outside_native_vector_levels_are_atomic() {
        var model = IWorkPagesDocument.Create(); var list = model.CreateNumberedList();
        for (int level = 0; level < 8; level++) list.AddItem("Parent " + level, level);
        list.AddItem("Deepest last", 8, int.MaxValue); byte[] before = model.SaveBytes();
        Assert.Throws<InvalidOperationException>(() => list.AddItem("Overflow", 8));
        Assert.Throws<ArgumentOutOfRangeException>(() => list.AddItem("Outside native vector", 9));
        Assert.Throws<ArgumentOutOfRangeException>(() => list.AddRichItem(new[] { new IWorkTextRunData("Outside", "Arial", 12) }, 9));
        Assert.Equal(before, model.SaveBytes());
        Assert.Equal(2, list.AddItem("Advance immediate parent", 7).ListNumber);
        Assert.Equal(1, list.AddRichItem(new[] { new IWorkTextRunData("Reset deepest", "Arial", 12) }, 8).ListNumber);
    }

    [Fact]
    public void Uniform_text_edits_at_every_native_level_preserve_all_counters_and_native_records() {
        var source = IWorkSourceDocument.Open(Example().SaveBytes()); var before = source.ReadPages().Body;
        var edit = source.BeginEdit();
        foreach (string text in LevelNames) {
            int index = before.Paragraphs.ToList().FindIndex(p => p.Text == text);
            edit.SetPagesBodyParagraphText(before.SourceIdentity!, index + 1, text.ToUpperInvariant());
        }
        byte[] pending = edit.SaveBytes(new() { AllowUnassessedVisualCaches = true });
        var saved = IWorkSourceDocument.Open(pending); var after = saved.ReadPages().Body;
        Assert.True(after.IsFormattingComplete);
        Assert.Equal(before.Paragraphs.Select(p => p.ListNumber), after.Paragraphs.Select(p => p.ListNumber));
        Assert.Equal(before.Paragraphs.Select(p => p.ListLevel), after.Paragraphs.Select(p => p.ListLevel));
        Assert.Equal(before.Paragraphs.Select(p => LevelNames.Contains(p.Text) ? p.Text.ToUpperInvariant() : p.Text),
            after.Paragraphs.Select(p => p.Text));
        foreach (var original in source.Records) {
            byte[] expected = original.GetPayload();
            if (original.Identifier == before.SourceIdentity!.RecordIdentifier && original.PayloadIndex == 0)
                foreach (string text in LevelNames) {
                    byte[] oldText = Encoding.UTF8.GetBytes(text); int at = expected.AsSpan().IndexOf(oldText);
                    Assert.True(at >= 0); Assert.Equal(-1, expected.AsSpan(at + oldText.Length).IndexOf(oldText));
                    Encoding.UTF8.GetBytes(text.ToUpperInvariant()).CopyTo(expected, at);
                }
            Assert.Equal(expected, saved.Records.Single(r => r.Identifier == original.Identifier && r.PayloadIndex == original.PayloadIndex).GetPayload());
        }
        IWorkPagesDecimalTextEditingTests.CheckWord(saved);
        int deepest = before.Paragraphs.ToList().FindIndex(p => p.Text == "Level eight") + 1;
        Assert.Throws<NotSupportedException>(() => edit.SetPagesBodyParagraphText(before.SourceIdentity!, deepest, "short"));
        using var cancel = new CancellationTokenSource(); cancel.Cancel();
        Assert.Throws<OperationCanceledException>(() => edit.SetPagesBodyParagraphText(before.SourceIdentity!, deepest, "Level eight", cancel.Token));
        Assert.Equal(pending, edit.SaveBytes(new() { AllowUnassessedVisualCaches = true }));
        foreach (string text in LevelNames) {
            int index = before.Paragraphs.ToList().FindIndex(p => p.Text == text);
            edit.SetPagesBodyParagraphText(before.SourceIdentity!, index + 1, text);
        }
        Assert.Empty(edit.SaveBytesResult().Report.UpdatedRecords);
        Assert.Equal(source.ComputePackageContentHash(), edit.SaveBytesResult().Report.OutputPackageContentHash);
        Assert.Equal(LevelNames, before.Paragraphs.Skip(1).Take(9).Select(p => p.Text));
    }

    internal static IWorkPagesDocument Example(IWorkDecimalListMarkerStyle markerStyle = IWorkDecimalListMarkerStyle.Period) {
        var document = IWorkPagesDocument.Create(); document.AddParagraph("Nine decimal levels");
        var list = document.CreateNumberedList(markerStyle);
        for (int level = 0; level < 9; level++) list.AddItem(LevelNames[level], level, level == 0 ? 7 : null);
        list.AddItem("Deep continued two", 8); document.AddParagraph("Plain gap");
        list.AddItem("Deep after gap three", 8); list.AddItem("Middle restart four", 4, 4);
        foreach (int level in new[] { 5, 6, 7, 8 }) list.AddItem("Reset level " + level, level);
        list.AddItem("Deep restart nine", 8, 9);
        document.CreateNumberedList().AddItem("Other root three", restartAt: 3);
        list.AddItem("Resumed deep ten", 8); list.AddItem("Root eight");
        for (int level = 1; level < 9; level++) list.AddItem("Root reset level " + level, level);
        list.AddItem("Root restart four", restartAt: 4);
        for (int level = 1; level < 9; level++) list.AddItem("Restart level " + level, level);
        document.AddParagraph("End nine-level profile"); return document;
    }
}
