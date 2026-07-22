using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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

        RtfConversionResult<WordDocument> result = RtfDocument
            .Read(rtf, RtfReadOptions.CreateUntrustedProfile())
            .ToWordDocumentResult("string");
        using (result.Value) {
            RtfConversionDiagnostic diagnostic = Assert.Single(result.Report.Diagnostics, item => item.Code == "RTF105");
            Assert.Equal(RtfConversionAction.Blocked, diagnostic.Action);
            Assert.Equal("string", diagnostic.SourcePath);
        }
    }

    [Fact]
    public void Word_Rtf_Read_Result_Uses_Bounded_Core_Profile() {
        var options = RtfReadOptions.CreateUntrustedProfile();
        options.MaxInputCharacters = 4;

        Assert.Throws<RtfReadLimitException>(() =>
            RtfDocument.Read(@"{\rtf1 Too large}", options).ToWordDocumentResult("string"));
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
        definition.AddLevel(RtfListKind.Decimal);
        definition.AddLevel(RtfListKind.Decimal);
        RtfListOverride listOverride = rtf.AddListOverride(20, 10);
        RtfListLevelOverride levelOverride = listOverride.AddLevelOverride();
        levelOverride.LevelIndex = 2;
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
        Assert.Equal(2, roundTripOverride.LevelIndex);
        Assert.True(roundTripOverride.OverrideStartAt);
        RtfListLevelOverride inactiveRoundTripOverride = Assert.Single(Assert.Single(roundTrip.ListOverrides, item => item.Id == 21).LevelOverrides);
        Assert.Null(inactiveRoundTripOverride.StartAt);
        Assert.False(inactiveRoundTripOverride.OverrideStartAt);
        Assert.Contains(toWord.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfWordStylesMapped");
        Assert.Contains(toWord.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfWordListDefinitionsMapped");
        toWord.Report.RequireNoLoss();
    }

    [Fact]
    public void Rtf_ToWord_Synthesizes_Numbering_For_Public_SetList_Paragraphs() {
        RtfDocument rtf = RtfDocument.Create();
        RtfParagraph paragraph = rtf.AddParagraph("Generated list item").SetList(listId: 7, level: 1, kind: RtfListKind.Bullet);

        RtfConversionResult<WordDocument> result = rtf.ToWordDocumentResult();
        using WordDocument word = result.Value;

        Numbering numbering = word._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        NumberingInstance instance = Assert.Single(numbering.Elements<NumberingInstance>());
        Assert.Equal(7, (int?)instance.NumberID);
        AbstractNum definition = Assert.Single(numbering.Elements<AbstractNum>());
        Assert.Contains(definition.Elements<Level>(), level => (int?)level.LevelIndex == 1 && level.NumberingFormat?.Val?.Value == NumberFormatValues.Bullet);
        Assert.Equal(7, Assert.Single(word.Paragraphs)._listNumberId);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfWordListDefinitionsMapped" && diagnostic.Count == 2);
        result.RequireNoLoss();
    }

    [Fact]
    public void Rtf_ToWord_Result_Reports_Object_And_Shape_Loss_In_Headers_And_Footers() {
        RtfDocument rtf = RtfDocument.Create();
        rtf.AddParagraph("Body");
        rtf.AddHeader().AddParagraph("Header").AddObject(RtfObjectKind.Embedded, new byte[] { 1 });
        rtf.AddFooter().AddParagraph("Footer").AddShape();

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
    public void Word_Bridge_Preserves_Sparse_Numbering_Level_And_Positive_First_Line_Indent() {
        using WordDocument word = WordDocument.Create();
        MainDocumentPart main = word._wordprocessingDocument.MainDocumentPart!;
        NumberingDefinitionsPart numberingPart = main.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering = new Numbering(
            new AbstractNum(
                new Level(
                    new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "%3." },
                    new PreviousParagraphProperties(
                        new Indentation { Left = "720", FirstLine = "240" })) {
                    LevelIndex = 2
                }) {
                AbstractNumberId = 10
            });

        RtfDocument rtf = word.ToRtfDocument();

        RtfListDefinition definition = Assert.Single(rtf.ListDefinitions);
        Assert.Equal(3, definition.Levels.Count);
        RtfListLevel level = definition.Levels[2];
        Assert.Equal(2, level.LevelIndex);
        Assert.Equal(240, level.FirstLineIndentTwips);
    }

    [Fact]
    public void Word_Bridge_Rejects_Numbering_Levels_Outside_Word_Range() {
        using WordDocument word = WordDocument.Create();
        MainDocumentPart main = word._wordprocessingDocument.MainDocumentPart!;
        NumberingDefinitionsPart numberingPart = main.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering = new Numbering(
            new AbstractNum(
                new Level(new NumberingFormat { Val = NumberFormatValues.Decimal }) { LevelIndex = 9 }) {
                AbstractNumberId = 10
            });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => word.ToRtfDocument());

        Assert.Contains("outside the supported range", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Rtf_Bridge_Maps_Large_Style_Sets_Without_Repeated_Scans() {
        RtfDocument rtf = RtfDocument.Create();
        for (int index = 0; index < 1_000; index++) rtf.AddStyle(index + 1, "Style " + index);

        using WordDocument word = rtf.ToWordDocument();

        Styles styles = word._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Assert.Equal(1_000, styles.Elements<Style>().Count(style => style.StyleId?.Value?.StartsWith("RtfP", StringComparison.Ordinal) == true));
    }

    [Fact]
    public void Word_ToRtf_Result_Reports_Unsupported_Content_Inside_Table_Cells() {
        using WordDocument word = WordDocument.Create();
        WordTable table = word.AddTable(1, 1);
        table.Rows[0].Cells[0]._tableCell.Append(new SdtBlock(
            new SdtContentBlock(new Paragraph(new Run(new Text("Controlled"))))));

        RtfConversionResult<RtfDocument> result = word.ToRtfDocumentResult();

        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "WordRtfElementOmitted" &&
            diagnostic.Feature == nameof(WordStructuredDocumentTag) &&
            diagnostic.Count == 1);
        Assert.Throws<RtfConversionLossException>(() => result.RequireNoLoss());
    }

    [Fact]
    public void Word_ToRtf_Result_Reports_Unsupported_Content_Inside_Paragraphs() {
        using WordDocument word = WordDocument.Create();
        word.AddParagraph().AddStructuredDocumentTag("Controlled");

        RtfConversionResult<RtfDocument> result = word.ToRtfDocumentResult();

        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "WordRtfElementOmitted" &&
            diagnostic.Feature == nameof(WordStructuredDocumentTag));
        Assert.Throws<RtfConversionLossException>(() => result.RequireNoLoss());
    }

    [Fact]
    public void Word_ToRtf_Result_Reports_Unsupported_Header_And_Footer_Elements() {
        using WordDocument word = WordDocument.Create();
        word.HeaderDefaultOrCreate._header!.Append(new SdtBlock(
            new SdtContentBlock(new Paragraph(new Run(new Text("Header control"))))));
        word.FooterDefaultOrCreate._footer!.Append(new SdtBlock(
            new SdtContentBlock(new Paragraph(new Run(new Text("Footer control"))))));

        RtfConversionResult<RtfDocument> result = word.ToRtfDocumentResult();

        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "WordRtfElementOmitted" &&
            diagnostic.Feature == nameof(WordStructuredDocumentTag) &&
            diagnostic.Count == 2);
        Assert.Throws<RtfConversionLossException>(() => result.RequireNoLoss());
    }
}
