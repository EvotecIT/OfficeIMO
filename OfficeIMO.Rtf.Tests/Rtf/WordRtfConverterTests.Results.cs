using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class WordRtfConverterTests {
    [Fact]
    public void Rtf_ToWord_Result_Reports_Mapped_Styles_And_Lists_With_Object_Shape_Loss() {
        RtfDocument rtf = RtfDocument.Create();
        rtf.AddStyle(7, "Clinical");
        rtf.AddListDefinition(10, "Steps");
        rtf.AddListOverride(20, 10);
        rtf.AddObject(RtfObjectKind.Embedded, new byte[] { 1, 2, 3 });
        rtf.AddShape().AddTextBoxParagraph("Shape text");

        RtfConversionResult<WordDocument> result = rtf.ToWordDocumentResult();
        using (result.Value) {
            Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfWordStylesMapped" && diagnostic.Action == RtfConversionAction.Preserved);
            Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfWordListDefinitionsMapped" && diagnostic.Count == 2);
            Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfWordObjectsOmitted" && diagnostic.Action == RtfConversionAction.Omitted);
            Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfWordShapesOmitted" && diagnostic.Action == RtfConversionAction.Omitted);
            Assert.Throws<RtfConversionLossException>(() => result.RequireNoLoss());
        }
    }

    [Fact]
    public void Rtf_ToWord_Result_Reports_Losses_Inside_Nested_Table_Cells() {
        RtfDocument rtf = RtfDocument.Create();
        RtfTable outer = rtf.AddTable(1, 1);
        RtfTable nested = outer.Rows[0].Cells[0].AddTable(1, 1);
        RtfParagraph nestedParagraph = nested.Rows[0].Cells[0].AddParagraph("Nested");
        nestedParagraph.AddObject(RtfObjectKind.Embedded, new byte[] { 1, 2, 3 });
        nestedParagraph.AddShape().AddTextBoxParagraph("Shape");

        RtfConversionResult<WordDocument> result = rtf.ToWordDocumentResult();
        using (result.Value) {
            Assert.Contains(result.Report.Diagnostics, diagnostic =>
                diagnostic.Code == "RtfWordObjectsOmitted" && diagnostic.Count == 1);
            Assert.Contains(result.Report.Diagnostics, diagnostic =>
                diagnostic.Code == "RtfWordShapesOmitted" && diagnostic.Count == 1);
            Assert.Throws<RtfConversionLossException>(() => result.RequireNoLoss());
        }
    }

    [Fact]
    public void Word_Rtf_Read_Result_Combines_Core_Policy_And_Bridge_Diagnostics() {
        const string rtf = @"{\rtf1{\object\objemb{\*\objdata 0102}}Visible}";

        RtfConversionResult<WordDocument> result = rtf.LoadFromRtfResult();
        using (result.Value) {
            RtfConversionDiagnostic diagnostic = Assert.Single(result.Report.Diagnostics, item => item.Code == "RTF105");
            Assert.Equal(RtfConversionAction.Blocked, diagnostic.Action);
            Assert.Equal("string", diagnostic.SourcePath);
        }
    }

    [Fact]
    public async Task Word_Rtf_Async_Read_Result_Uses_Bounded_Core_Profile() {
        var options = RtfReadOptions.CreateUntrustedProfile();
        options.MaxInputCharacters = 4;

        await Assert.ThrowsAsync<RtfReadLimitException>(() =>
            @"{\rtf1 Too large}".LoadFromRtfResultAsync(options));
    }

    [Fact]
    public void Word_Bridge_Preserves_Style_And_Numbering_Structures_In_Both_Directions() {
        RtfDocument rtf = RtfDocument.Create();
        RtfStyle paragraphStyle = rtf.AddStyle(7, "Clinical", RtfStyleKind.Paragraph);
        paragraphStyle.Bold = true;
        paragraphStyle.SpaceAfterTwips = 120;
        RtfStyle characterStyle = rtf.AddStyle(8, "Emphasis", RtfStyleKind.Character);
        characterStyle.Italic = true;
        RtfListDefinition definition = rtf.AddListDefinition(10, "Steps");
        RtfListLevel level = definition.AddLevel(RtfListKind.Decimal);
        level.StartAt = 3;
        level.Text = "%1.";
        level.LeftIndentTwips = 720;
        RtfListOverride listOverride = rtf.AddListOverride(20, 10);
        RtfListLevelOverride levelOverride = listOverride.AddLevelOverride();
        levelOverride.StartAt = 7;
        levelOverride.OverrideStartAt = true;
        RtfListLevelOverride inactiveOverride = rtf.AddListOverride(21, 10).AddLevelOverride();
        inactiveOverride.StartAt = 9;
        inactiveOverride.OverrideStartAt = false;
        RtfParagraph paragraph = rtf.AddParagraph().SetStyle(7).SetList(20, 0, RtfListKind.Decimal);
        paragraph.ListDefinitionId = 10;
        paragraph.AddText("Styled list item").SetStyle(8);

        RtfConversionResult<WordDocument> toWord = rtf.ToWordDocumentResult();
        using WordDocument word = toWord.Value;
        WordParagraph wordParagraph = Assert.Single(word.Paragraphs);
        Assert.Equal("RtfP7", wordParagraph.StyleId);
        Assert.True(wordParagraph.IsListItem);
        Assert.Equal(20, wordParagraph._listNumberId);
        Assert.Equal("RtfC8", Assert.Single(wordParagraph.GetRuns()).CharacterStyleId);

        RtfConversionResult<RtfDocument> roundTripResult = word.ToRtfDocumentResult();
        RtfDocument roundTrip = roundTripResult.Value;
        RtfParagraph roundTripParagraph = Assert.Single(roundTrip.Paragraphs);
        RtfStyle mappedParagraphStyle = Assert.Single(roundTrip.Styles, style => style.Id == roundTripParagraph.StyleId && style.Kind == RtfStyleKind.Paragraph);
        RtfRun roundTripRun = Assert.Single(roundTripParagraph.Runs);
        RtfStyle mappedCharacterStyle = Assert.Single(roundTrip.Styles, style => style.Id == roundTripRun.StyleId && style.Kind == RtfStyleKind.Character);

        Assert.Equal("RtfP7", mappedParagraphStyle.Name);
        Assert.Equal("RtfC8", mappedCharacterStyle.Name);
        Assert.Equal(20, roundTripParagraph.ListId);
        Assert.Equal(10, roundTripParagraph.ListDefinitionId);
        Assert.Equal(3, Assert.Single(roundTrip.ListDefinitions, item => item.Id == 10).Levels[0].StartAt);
        RtfListLevelOverride roundTripOverride = Assert.Single(Assert.Single(roundTrip.ListOverrides, item => item.Id == 20).LevelOverrides);
        Assert.Equal(7, roundTripOverride.StartAt);
        Assert.True(roundTripOverride.OverrideStartAt);
        RtfListLevelOverride inactiveRoundTripOverride = Assert.Single(Assert.Single(roundTrip.ListOverrides, item => item.Id == 21).LevelOverrides);
        Assert.Null(inactiveRoundTripOverride.StartAt);
        Assert.False(inactiveRoundTripOverride.OverrideStartAt);
        Assert.Contains(toWord.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfWordStylesMapped");
        Assert.Contains(toWord.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfWordListDefinitionsMapped");
        toWord.Report.RequireNoLoss();
    }
}
