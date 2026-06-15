using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlListTableTests {
    [Fact]
    public void RtfDocument_ToHtml_RoundTrips_Rich_List_Table_Metadata() {
        RtfDocument document = RtfDocument.Create();
        RtfListDefinition definition = document.AddListDefinition(100, "Rich");
        definition.TemplateId = 77;
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

        RtfListOverride listOverride = document.AddListOverride(3, 100);
        listOverride.OverrideCount = 2;
        RtfListLevelOverride firstOverride = listOverride.AddLevelOverride();
        firstOverride.OverrideFormat = true;
        firstOverride.OverrideStartAt = true;
        firstOverride.StartAt = 9;
        RtfListLevelOverride secondOverride = listOverride.AddLevelOverride();
        secondOverride.OverrideStartAt = false;

        document.AddParagraph("Item").SetList(listId: 3, level: 0, kind: RtfListKind.Decimal);

        string html = document.ToHtml(new RtfHtmlSaveOptions {
            FragmentOnly = false,
            NewLine = "\n"
        });

        Assert.Contains("<meta name=\"officeimo-rtf-lists\" content=\"", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.ToRtfDocumentFromHtml();

        RtfListDefinition roundTripDefinition = Assert.Single(roundTrip.ListDefinitions);
        Assert.Equal(100, roundTripDefinition.Id);
        Assert.Equal("Rich", roundTripDefinition.Name);
        Assert.Equal(77, roundTripDefinition.TemplateId);

        RtfListLevel roundTripLevel = Assert.Single(roundTripDefinition.Levels);
        Assert.Equal(RtfListKind.Decimal, roundTripLevel.Kind);
        Assert.Equal(0, roundTripLevel.NumberFormat);
        Assert.Equal(2, roundTripLevel.NumberFormatN);
        Assert.Equal(RtfListLevelAlignment.Right, roundTripLevel.Alignment);
        Assert.Equal(RtfListLevelAlignment.Center, roundTripLevel.AlignmentN);
        Assert.Equal(RtfListLevelFollowCharacter.Space, roundTripLevel.FollowCharacter);
        Assert.Equal(7, roundTripLevel.StartAt);
        Assert.Equal(120, roundTripLevel.SpaceTwips);
        Assert.Equal(240, roundTripLevel.IndentTwips);
        Assert.True(roundTripLevel.LegalNumbering);
        Assert.True(roundTripLevel.NoRestart);
        Assert.Equal(3, roundTripLevel.PictureIndex);
        Assert.True(roundTripLevel.PictureNoSize);
        Assert.Equal("%1.", roundTripLevel.Text);
        Assert.Equal("\u0001", roundTripLevel.Numbers);
        Assert.Equal(1080, roundTripLevel.LeftIndentTwips);
        Assert.Equal(-360, roundTripLevel.FirstLineIndentTwips);

        RtfListOverride roundTripOverride = Assert.Single(roundTrip.ListOverrides);
        Assert.Equal(3, roundTripOverride.Id);
        Assert.Equal(100, roundTripOverride.ListId);
        Assert.Equal(2, roundTripOverride.OverrideCount);
        Assert.Collection(roundTripOverride.LevelOverrides,
            item => {
                Assert.True(item.OverrideFormat);
                Assert.True(item.OverrideStartAt);
                Assert.Equal(9, item.StartAt);
            },
            item => {
                Assert.Null(item.OverrideFormat);
                Assert.False(item.OverrideStartAt);
                Assert.Null(item.StartAt);
            });

        RtfParagraph paragraph = Assert.Single(roundTrip.Paragraphs);
        Assert.Equal(3, paragraph.ListId);
        Assert.Equal(100, paragraph.ListDefinitionId);

        string rtf = roundTrip.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"\listtemplateid77", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\levelnfc0\levelnfcn2\leveljc2\leveljcn1\levelfollow1\levelstartat7\levelspace120\levelindent240\levellegal1\levelnorestart1\levelpicture3\levelpicturenosize", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\*\listoverridetable{\listoverride\listid100\listoverridecount2{\lfolevel\listoverrideformat1\listoverridestartat1\levelstartat9}{\lfolevel\listoverridestartat0}\ls3}}", rtf, StringComparison.Ordinal);
    }
}
