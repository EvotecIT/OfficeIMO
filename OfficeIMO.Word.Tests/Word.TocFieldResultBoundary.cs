using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordTocFieldResultBoundaryTests {
    [Theory]
    [InlineData("headings", 0)]
    [InlineData("headings", 1)]
    [InlineData("figures", 0)]
    [InlineData("figures", 1)]
    [InlineData("index", 0)]
    [InlineData("index", 1)]
    public void CustomEmptyMessageSurvivesRepeatedRefreshAndReopen(string kind, int headings) {
        using var output = new MemoryStream();
        using (WordDocument document = WordDocument.Create()) {
            var contents = document.AddTableOfContent();
            for (int index = 0; index < headings; index++) {
                document.AddParagraph("Chapter").SetStyle(WordParagraphStyles.Heading1);
                if (kind == "figures") document.AddParagraph("Figure ")._paragraph.Append(
                    new SimpleField(new Run(new Text("1"))) { Instruction = " SEQ Figure " }, new Run(new Text(" Chapter")));
                if (kind == "index") document.AddParagraph("Chapter ")._paragraph.Append(
                    new SimpleField(new Run(new Text(string.Empty))) { Instruction = " XE \"Chapter\" " });
            }
            Action refresh = () => {
                if (kind == "figures") contents.RefreshListOfFigures();
                else if (kind == "index") contents.RefreshIndex();
                else contents.RefreshEntries();
            };
            refresh();
            contents.TextNoContent = "Nothing to list yet";
            if (headings == 0) Assert.Contains("Nothing to list yet", Assert.Single(document.InspectFields(), item => item.FieldType == (kind == "index" ? WordFieldType.Index : WordFieldType.TOC)).ResultText);
            refresh();
            refresh();
            Assert.Equal("Nothing to list yet", contents.TextNoContent);
            contents.TextNoContent = "No chapters available";
            document.Save(output);
        }
        output.Position = 0;
        using var reopened = WordDocument.Load(output);
        Assert.Equal("No chapters available", reopened.TableOfContent!.TextNoContent);
        if (kind == "figures") reopened.TableOfContent.RefreshListOfFigures();
        else if (kind == "index") reopened.TableOfContent.RefreshIndex();
        else reopened.TableOfContent.RefreshEntries();
        var field = Assert.Single(reopened.InspectFields(), item => item.FieldType == (kind == "index" ? WordFieldType.Index : WordFieldType.TOC));
        Assert.Contains(headings == 0 ? "No chapters available" : "Chapter", field.ResultText);
        Assert.True(reopened.DocumentIsValid, string.Join("; ", reopened.DocumentValidationErrors));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void RepeatedRefreshAndLevelChangesPreserveNestedPageReferences(int headings) {
        using var document = WordDocument.Create();
        var contents = document.AddTableOfContent();
        for (int index = 0; index < headings; index++) {
            document.AddParagraph("Chapter " + index).SetStyle(WordParagraphStyles.Heading1);
        }
        contents.RefreshEntries();
        string instruction = Assert.Single(document.InspectFields(), field => field.FieldType == WordFieldType.TOC).InstructionText;
        for (int pass = 0; pass < 3; pass++) {
            contents.RefreshEntries();
            Assert.Equal(instruction, Assert.Single(document.InspectFields(), field => field.FieldType == WordFieldType.TOC).InstructionText);
        }
        string[] pageReferences = document.InspectFields().Where(field => field.FieldType == WordFieldType.PageRef)
            .Select(field => field.InstructionText).ToArray();
        Assert.Equal(headings, pageReferences.Length);
        contents.SetLevels(1, 2);
        Assert.Equal(pageReferences, document.InspectFields().Where(field => field.FieldType == WordFieldType.PageRef)
            .Select(field => field.InstructionText).ToArray());
        Assert.Contains("\\o \"1-2\"", Assert.Single(document.InspectFields(), field => field.FieldType == WordFieldType.TOC).InstructionText);
    }

    [Theory]
    [InlineData("save", "Summary", "Evidence")]
    [InlineData("headings", "Summary", "Evidence")]
    [InlineData("figures", "Boundary caption", "")]
    [InlineData("index", "Boundary topic", "")]
    public void SavedGeneratedEntriesBelongToTheFieldResult(string kind, string first, string second) {
        using var output = new MemoryStream();
        using (WordDocument document = WordDocument.Create()) {
            WordTableOfContent contents = document.AddTableOfContent(minLevel: 1, maxLevel: 1);
            contents.Text = "Contents";
            document.AddParagraph("Summary").SetStyle(WordParagraphStyles.Heading1);
            document.AddPageBreak();
            document.AddParagraph("Evidence").SetStyle(WordParagraphStyles.Heading1);
            if (kind == "headings") { contents.RefreshEntries(); }
            if (kind == "figures") {
                document.AddParagraph("Figure ")._paragraph.Append(
                    new SimpleField(new Run(new Text("1"))) { Instruction = " SEQ Figure " },
                    new Run(new Text(" Boundary caption")));
                contents.RefreshListOfFigures();
            }
            if (kind == "index") {
                document.AddParagraph("Topic ")._paragraph.Append(
                    new SimpleField(new Run(new Text(string.Empty))) { Instruction = " XE \"Boundary topic\" " });
                contents.RefreshIndex();
            }
            document.Save(output);
        }
        output.Position = 0;
        using WordDocument reopened = WordDocument.Load(output);
        Assert.NotEmpty(reopened.TableOfContent!.TextNoContent);
        WordFieldInfo field = Assert.Single(reopened.InspectFields(), item =>
            item.InstructionText.TrimStart().StartsWith(kind == "index" ? "INDEX" : "TOC", StringComparison.Ordinal));
        foreach (string entry in new[] { first, second }.Where(value => value.Length > 0)) {
            Assert.Contains(entry, field.ResultText);
            Assert.Equal(1, field.ResultText.Split(new[] { entry }, StringSplitOptions.None).Length - 1);
        }
        Assert.True(reopened.DocumentIsValid, string.Join("; ", reopened.DocumentValidationErrors));
    }
}
