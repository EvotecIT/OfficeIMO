using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class RtfDocumentRichFeatureTests {
    [Fact]
    public void Write_And_Read_Explicit_List_Definition_And_Override() {
        RtfDocument document = RtfDocument.Create();
        RtfListDefinition definition = document.AddListDefinition(100, "Decimal");
        RtfListLevel level = definition.AddLevel(RtfListKind.Decimal);
        level.Text = "%1.";
        level.Numbers = "\u0001";
        level.LeftIndentTwips = 720;
        level.FirstLineIndentTwips = -360;
        document.AddListOverride(3, 100);
        document.AddParagraph("First").SetList(listId: 3, level: 0, kind: RtfListKind.Decimal);

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\*\listtable", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\listid100", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\*\listoverridetable{\listoverride\listid100\listoverridecount0\ls3}}", rtf, StringComparison.Ordinal);
        RtfListDefinition readDefinition = Assert.Single(read.Document.ListDefinitions);
        Assert.Equal(100, readDefinition.Id);
        Assert.Equal("Decimal", readDefinition.Name);
        RtfListOverride readOverride = Assert.Single(read.Document.ListOverrides);
        Assert.Equal(3, readOverride.Id);
        Assert.Equal(100, readOverride.ListId);
        RtfParagraph paragraph = Assert.Single(read.Document.Paragraphs);
        Assert.Equal("First", paragraph.ToPlainText());
        Assert.Equal(3, paragraph.ListId);
        Assert.Equal(100, paragraph.ListDefinitionId);
        Assert.Equal(RtfListKind.Decimal, paragraph.ListKind);
    }

    [Fact]
    public void Write_And_Read_Rich_List_Level_Metadata() {
        RtfDocument document = RtfDocument.Create();
        RtfListDefinition definition = document.AddListDefinition(100, "Rich");
        RtfListLevel level = definition.AddLevel(RtfListKind.Decimal);
        level.NumberFormat = 0;
        level.NumberFormatN = 2;
        level.Alignment = RtfListLevelAlignment.Right;
        level.AlignmentN = RtfListLevelAlignment.Center;
        level.FollowCharacter = RtfListLevelFollowCharacter.Space;
        level.StartAt = 7;
        level.SpaceTwips = 120;
        level.IndentTwips = 240;
        level.LegalNumbering = true;
        level.NoRestart = true;
        level.PictureIndex = 3;
        level.PictureNoSize = true;
        level.Text = "%1.";
        level.Numbers = "\u0001";
        level.LeftIndentTwips = 1080;
        level.FirstLineIndentTwips = -360;
        document.AddListOverride(3, 100);
        document.AddParagraph("Item").SetList(listId: 3, level: 0, kind: RtfListKind.Decimal);

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"\levelnfc0\levelnfcn2\leveljc2\leveljcn1\levelfollow1\levelstartat7\levelspace120\levelindent240\levellegal1\levelnorestart1\levelpicture3\levelpicturenosize", rtf, StringComparison.Ordinal);
        RtfListLevel readLevel = Assert.Single(Assert.Single(read.Document.ListDefinitions).Levels);
        Assert.Equal(0, readLevel.NumberFormat);
        Assert.Equal(2, readLevel.NumberFormatN);
        Assert.Equal(RtfListLevelAlignment.Right, readLevel.Alignment);
        Assert.Equal(RtfListLevelAlignment.Center, readLevel.AlignmentN);
        Assert.Equal(RtfListLevelFollowCharacter.Space, readLevel.FollowCharacter);
        Assert.Equal(7, readLevel.StartAt);
        Assert.Equal(120, readLevel.SpaceTwips);
        Assert.Equal(240, readLevel.IndentTwips);
        Assert.True(readLevel.LegalNumbering);
        Assert.True(readLevel.NoRestart);
        Assert.Equal(3, readLevel.PictureIndex);
        Assert.True(readLevel.PictureNoSize);
        Assert.Equal(1080, readLevel.LeftIndentTwips);
        Assert.Equal(-360, readLevel.FirstLineIndentTwips);
    }

    [Fact]
    public void Write_And_Read_List_Level_Overrides() {
        RtfDocument document = RtfDocument.Create();
        RtfListDefinition definition = document.AddListDefinition(100, "Decimal");
        RtfListLevel level = definition.AddLevel(RtfListKind.Decimal);
        level.Text = "%1.";
        level.Numbers = "\u0001";
        level.LeftIndentTwips = 720;
        level.FirstLineIndentTwips = -360;
        RtfListOverride listOverride = document.AddListOverride(3, 100);
        RtfListLevelOverride firstOverride = listOverride.AddLevelOverride();
        firstOverride.OverrideFormat = true;
        firstOverride.OverrideStartAt = true;
        firstOverride.StartAt = 9;
        RtfListLevelOverride secondOverride = listOverride.AddLevelOverride();
        secondOverride.OverrideStartAt = false;
        document.AddParagraph("Item").SetList(listId: 3, level: 0, kind: RtfListKind.Decimal);

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\*\listoverridetable{\listoverride\listid100\listoverridecount2{\lfolevel\listoverrideformat1\listoverridestartat1\levelstartat9}{\lfolevel\listoverridestartat0}\ls3}}", rtf, StringComparison.Ordinal);
        RtfListOverride readOverride = Assert.Single(read.Document.ListOverrides);
        Assert.Equal(2, readOverride.OverrideCount);
        Assert.Collection(readOverride.LevelOverrides,
            levelOverride => {
                Assert.True(levelOverride.OverrideFormat);
                Assert.True(levelOverride.OverrideStartAt);
                Assert.Equal(9, levelOverride.StartAt);
            },
            levelOverride => {
                Assert.Null(levelOverride.OverrideFormat);
                Assert.False(levelOverride.OverrideStartAt);
                Assert.Null(levelOverride.StartAt);
            });
    }
}
