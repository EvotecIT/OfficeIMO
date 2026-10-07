using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(-1, false, "1,1,1,1")]
    [InlineData(0, false, "1,2,3,4")]
    [InlineData(1, false, "1,2,1,2")]
    [InlineData(0, true, "1,1,1,1")]
    public void ListRestartRulesFollowWordAcrossParentsAndSkippedLevels(int restart, bool instanceOverride, string childIndices) {
        using WordDocument document = CreateListInstanceCounterControl(out WordParagraph[] items, out Numbering numbering, out NumberingInstance second, 9);
        AbstractNum abstractNum = numbering.Elements<AbstractNum>().Single();
        foreach (Level level in abstractNum.Elements<Level>()) {
            if (level.LevelIndex!.Value != 0) level.StartNumberingValue!.Val = 1;
            level.LevelText!.Val = string.Join(".", Enumerable.Range(1, level.LevelIndex.Value + 1).Select(index => "%" + index)) + ".";
            level.NumberingFormat!.Val = NumberFormatValues.Decimal;
        }
        Level child = abstractNum.Elements<Level>().Single(level => level.LevelIndex!.Value == 2);
        if (restart >= 0) {
            if (instanceOverride) {
                child = (Level)child.CloneNode(true);
                second.Append(new LevelOverride(child) { LevelIndex = 2 });
            }
            child.AddChild(new LevelRestart { Val = restart }, true);
        }
        int[] levels = { 0, 1, 2, 1, 2, 0, 2, 1, 2 };
        for (int index = 0; index < items.Length; index++) {
            NumberingProperties properties = items[index]._paragraph.ParagraphProperties!.NumberingProperties!;
            properties.NumberingId!.Val = instanceOverride ? 2 : 1;
            properties.NumberingLevelReference!.Val = levels[index];
        }
        string[] values = childIndices.Split(',');
        string[] expected = { "12.", "12.1.", "12.1." + values[0] + ".", "12.2.", "12.2." + values[1] + ".",
            "13.", "13.1." + values[2] + ".", "13.2.", "13.2." + values[3] + "." };
        string before = numbering.OuterXml;
        var markers = WordDocumentTraversal.BuildListMarkers(document);
        Assert.Equal(expected, items.Select(item => markers[item].Marker));
        Assert.Equal(before, numbering.OuterXml);
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NestedListMarkersUseTheCurrentInstancesParentFormat(bool returnToFirstInstance) {
        using WordDocument document = CreateListInstanceCounterControl(out WordParagraph[] items, out Numbering numbering, out NumberingInstance second);
        AbstractNum abstractNum = numbering.Elements<AbstractNum>().Single();
        Level parentOverride = (Level)abstractNum.Elements<Level>().First().CloneNode(true);
        parentOverride.NumberingFormat!.Val = NumberFormatValues.LowerRoman;
        second.Append(new LevelOverride(parentOverride) { LevelIndex = 0 });
        Level child = abstractNum.Elements<Level>().Single(level => level.LevelIndex!.Value == 1);
        child.StartNumberingValue!.Val = 1; child.LevelText!.Val = "%1.%2.";
        child.NumberingFormat!.Val = NumberFormatValues.Decimal;
        for (int index = 1; index < items.Length; index++) {
            NumberingProperties properties = items[index]._paragraph.ParagraphProperties!.NumberingProperties!;
            properties.NumberingLevelReference!.Val = 1;
            properties.NumberingId!.Val = returnToFirstInstance && index == 3 ? 1 : 2;
        }
        var markers = WordDocumentTraversal.BuildListMarkers(document);
        Assert.Equal(new[] { "12.", "xii.1.", "xii.2.", returnToFirstInstance ? "12.3." : "xii.3." }, items.Select(item => markers[item].Marker));
        Assert.Empty(document.ValidateDocument());
    }
}
