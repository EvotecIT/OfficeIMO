using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfListStoryResourceRegressionTests {
    [Fact]
    public void ImportedStoryListsCannotCaptureImplicitDestinationLists() {
        RtfDocument destination = RtfDocument.Create();
        RtfParagraph original = destination.AddParagraph("Destination bullet").SetList();
        destination.AddNote(RtfNoteKind.Footnote).AddParagraph("Destination note").SetList(7);
        RtfDocument source = RtfDocument.Create();
        source.AddNote(RtfNoteKind.Footnote).AddParagraph("Imported decimal").SetList(3, kind: RtfListKind.Decimal);

        destination.AppendDocument(source).Report.RequireNoLoss();
        Assert.Equal(RtfListKind.Bullet, destination.ResolveListFormatting(original)!.Level.Kind);
        RtfParagraph imported = destination.Notes[1].Paragraphs[0];
        Assert.True(imported.ListId > 7);
        Assert.Equal(RtfListKind.Decimal, destination.ResolveListFormatting(imported)!.Level.Kind);
        RtfDocument reopened = RtfDocument.Read(destination.ToRtf()).Document;
        Assert.Equal(RtfListKind.Bullet, reopened.ResolveListFormatting(reopened.Paragraphs[0])!.Level.Kind);
        Assert.Equal(RtfListKind.Bullet, reopened.ResolveListFormatting(reopened.Notes[0].Paragraphs[0])!.Level.Kind);
        Assert.Equal(RtfListKind.Decimal, reopened.ResolveListFormatting(reopened.Notes[1].Paragraphs[0])!.Level.Kind);
    }

    [Fact]
    public void ListsInDetachedAndFieldReferencedNotesHaveIndependentSerializedResources() {
        RtfDocument source = RtfDocument.Create();
        source.AddNote(RtfNoteKind.Footnote).AddParagraph("Detached").SetList(3);
        var referenced = new RtfNote(RtfNoteKind.Footnote);
        referenced.AddParagraph("Referenced").SetList(4, kind: RtfListKind.Decimal);
        RtfRun run = source.AddParagraph("Anchor").AddField("QUOTE").Result.AddText("Field");
        run.Note = referenced;
        source.AddParagraph("Shared reference").Runs[0].Note = referenced;

        RtfDocument reopened = RtfDocument.Read(source.ToRtf()).Document;
        Assert.Equal(new[] { 3, 4 }, reopened.ListOverrides.Select(item => item.Id).OrderBy(id => id));
        Assert.Equal(2, reopened.ListDefinitions.Count);
        foreach (RtfParagraph paragraph in reopened.Notes.SelectMany(note => note.Paragraphs)) Assert.True(reopened.ResolveListFormatting(paragraph)!.IsDefined);
        Assert.Empty(source.ListOverrides);
        Assert.Empty(source.ListDefinitions);
    }

    [Fact]
    public void AppendingANoteOnlyListDoesNotBindItToTheDestinationsListWithTheSameId() {
        RtfDocument destination = RtfDocument.Create();
        destination.AddListDefinition(100).AddLevel(RtfListKind.Decimal).StartAt = 90;
        destination.AddListOverride(3, 100);
        destination.AddParagraph("Destination").SetList(3, kind: RtfListKind.Decimal);
        RtfDocument source = RtfDocument.Create();
        source.AddNote(RtfNoteKind.Footnote).AddParagraph("Imported bullet").SetList(3);

        destination.AppendDocument(source).Report.RequireNoLoss();
        RtfParagraph imported = Assert.Single(destination.Notes).Paragraphs[0];
        RtfListFormatting list = destination.ResolveListFormatting(imported)!;
        Assert.True(list.IsDefined);
        Assert.Equal(RtfListKind.Bullet, list.Level.Kind);
        Assert.NotEqual(3, list.InstanceId);
        Assert.NotEqual(100, list.DefinitionId);
        Assert.Equal(90, destination.ResolveListFormatting(destination.Paragraphs[0])!.Level.StartAt);
        Assert.Equal(3, source.Notes[0].Paragraphs[0].ListId);
    }

    [Fact]
    public void TextBoxListsAreDiscoveredThroughTableAndFieldContent() {
        RtfDocument source = RtfDocument.Create();
        RtfParagraph field = source.AddTable(1, 1).Rows[0].Cells[0].AddParagraph("Anchor").AddField("QUOTE").Result;
        field.AddShape().AddTextBoxParagraph("Text box").SetList(7, kind: RtfListKind.Decimal);
        RtfDocument reopened = RtfDocument.Read(source.ToRtf()).Document;
        Assert.Equal(7, Assert.Single(reopened.ListOverrides).Id);
        Assert.Equal(RtfListKind.Decimal, Assert.Single(reopened.ListDefinitions).Levels[0].Kind);
    }
}
