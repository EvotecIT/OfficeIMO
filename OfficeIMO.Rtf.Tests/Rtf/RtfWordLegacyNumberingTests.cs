using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfWordLegacyNumberingTests {
    [Fact]
    public void RichEdit_Numbering_Writes_Editable_Docx_Markers_Without_Mutating_Rtf() {
        const string item = @"{\*\pn\pnlvlbody\pnf0\pnindent360\pnstart1\pndec{\pntxta.}}\pard\fi-360\li720 ";
        const string input = @"{\rtf1\ansi\deff0{\fonttbl{\f0 Times New Roman;}}\f0\fs24"
            + @"{\pntext\f0 1.\tab}" + item + @"Alpha.\par"
            + @"{\pntext\f0 2.\tab}" + item + @"Bravo.\par}";
        RtfDocument source = RtfDocument.Parse(input);
        string before = Native(source);

        RtfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.False(conversion.Report.HasLoss);
        conversion.Report.RequireNoLoss();
        int firstId = Id(word.Paragraphs[0]);
        Assert.Equal(firstId, Id(word.Paragraphs[1]));
        Assert.Equal(new[] { "Alpha.", "Bravo." }, word.Paragraphs.Select(paragraph => paragraph.Text));
        Level level = LevelFor(word, firstId);
        Assert.Equal(NumberFormatValues.Decimal, level.NumberingFormat!.Val!.Value);
        Assert.Equal("%1.", level.LevelText!.Val!.Value);
        Assert.Equal(1, level.StartNumberingValue!.Val!.Value);
        Assert.Equal("Times New Roman", level.NumberingSymbolRunProperties!.RunFonts!.Ascii!.Value);
        AssertArtifact(word, "1.", "2.");
        using WordDocument repeated = source.ToWordDocument();
        AssertArtifact(repeated, "1.", "2.");
        Assert.Equal(before, Native(source));
        Assert.All(source.Paragraphs, paragraph => Assert.Null(paragraph.ListId));
    }

    [Fact]
    public void Anonymous_Legacy_Starts_Continue_And_Reset_After_An_Ordinary_Paragraph() {
        RtfDocument source = RtfDocument.Create();
        Legacy(source.AddParagraph("First"), 3);
        Legacy(source.AddParagraph("Second"), 1);
        source.AddParagraph("Between");
        Legacy(source.AddParagraph("Restarted"), 1);
        string before = Native(source);

        RtfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.False(conversion.Report.HasLoss);
        Assert.Equal(Id(word.Paragraphs[0]), Id(word.Paragraphs[1]));
        Assert.NotEqual(Id(word.Paragraphs[0]), Id(word.Paragraphs[3]));
        Assert.Equal(3, LevelFor(word, Id(word.Paragraphs[0])).StartNumberingValue!.Val!.Value);
        AssertArtifact(word, "3.", "4.", null, "1.");
        Assert.Equal(before, Native(source));
    }

    [Fact]
    public void Legacy_Bullets_And_Existing_Modern_Numbering_Keep_Their_Own_Resources() {
        RtfDocument source = RtfDocument.Create();
        RtfListLevel modern = source.AddListDefinition(100).AddLevel(RtfListKind.Decimal);
        modern.StartAt = 8;
        modern.Text = "%1.";
        source.AddListOverride(7, 100);
        source.AddParagraph("Modern first").SetList(7, kind: RtfListKind.Decimal);
        RtfParagraph bullet = source.AddParagraph("Bullet");
        bullet.SetLegacyNumbering(numbering => {
            numbering.LevelKind = RtfLegacyNumberingLevelKind.Bullet;
            numbering.TextBefore = "•";
        });
        bullet.SetListText("•\t");
        Assert.Equal(RtfListKind.Bullet, source.ResolveListFormatting(bullet)!.Level.Kind);
        source.AddParagraph("Modern continued").SetList(7, kind: RtfListKind.Decimal);
        string before = Native(source);

        RtfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.False(conversion.Report.HasLoss);
        Assert.Equal(7, Id(word.Paragraphs[0]));
        Assert.Equal(7, Id(word.Paragraphs[2]));
        Assert.NotEqual(7, Id(word.Paragraphs[1]));
        Assert.Equal(NumberFormatValues.Bullet, LevelFor(word, Id(word.Paragraphs[1])).NumberingFormat!.Val!.Value);
        AssertArtifact(word, "8.", "•", "9.");
        Assert.Equal(before, Native(source));
    }

    [Fact]
    public void Header_And_Table_Paragraphs_Use_The_Same_Projection_With_Independent_Story_Starts() {
        RtfDocument source = RtfDocument.Create();
        Legacy(source.AddParagraph("Body"), 3);
        RtfTable table = source.AddTable(1, 1);
        table.Rows[0].CellGapTwips = 0;
        Legacy(table.Rows[0].Cells[0].AddParagraph("Cell"), 1);
        Legacy(source.AddHeader().AddParagraph("Header"), 1);
        string before = Native(source);

        RtfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.False(conversion.Report.HasLoss);
        WordParagraph cell = Assert.Single(word.Tables[0].Rows[0].Cells[0].Paragraphs, paragraph => paragraph.Text == "Cell");
        WordParagraph header = Assert.Single(word.Header!.Default!.Paragraphs, paragraph => paragraph.Text == "Header");
        Assert.Equal(Id(word.Paragraphs[0]), Id(cell));
        Assert.NotEqual(Id(cell), Id(header));
        Assert.Equal(3, LevelFor(word, Id(cell)).StartNumberingValue!.Val!.Value);
        Assert.Equal(1, LevelFor(word, Id(header)).StartNumberingValue!.Val!.Value);
        Assert.Empty(new OpenXmlValidator().Validate(word._wordprocessingDocument));
        using var bytes = new MemoryStream();
        word.Save(bytes);
        bytes.Position = 0;
        using WordDocument reopened = WordDocument.Load(bytes);
        Assert.Equal(Id(reopened.Paragraphs[0]), Id(Assert.Single(reopened.Tables[0].Rows[0].Cells[0].Paragraphs, paragraph => paragraph.Text == "Cell")));
        Assert.Equal(before, Native(source));
    }

    [Theory]
    [InlineData("across-rows")]
    [InlineData("higher-level")]
    [InlineData("continuation")]
    [InlineData("marker-spacing")]
    [InlineData("cached-separator")]
    [InlineData("roman-resolved")]
    public void Unsupported_Legacy_Numbering_Preserves_Marker_Text_And_Reports_Flattening(string profile) {
        RtfDocument source = RtfDocument.Create();
        RtfParagraph paragraph = source.AddParagraph("Content");
        Legacy(paragraph, 1);
        string expected = "1.\t";
        switch (profile) {
            case "across-rows": paragraph.LegacyNumbering.NumberAcrossRows = true; break;
            case "higher-level": paragraph.LegacyNumbering.Level = 2; break;
            case "continuation": paragraph.LegacyNumbering.LevelKind = RtfLegacyNumberingLevelKind.Continue; break;
            case "marker-spacing": paragraph.LegacyNumbering.SpaceTwips = 240; break;
            case "cached-separator": expected = "1. "; break;
            case "roman-resolved":
                paragraph.LegacyNumbering.NumberStyle = RtfLegacyNumberingStyle.LowerRoman;
                paragraph.LegacyNumbering.StartAt = 9;
                paragraph.LegacyNumbering.TextBefore = "(";
                paragraph.LegacyNumbering.TextAfter = ")";
                expected = "(ix)\t";
                break;
        }
        if (profile != "roman-resolved") paragraph.SetListText(expected);
        string before = Native(source);

        RtfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        WordParagraphInfo visible = Assert.Single(word.Inspect().Sections
            .SelectMany(section => section.Elements).OfType<WordParagraphInfo>());
        Assert.Equal(expected + "Content", visible.Text);
        Assert.All(word.Paragraphs, item => Assert.Null(item._paragraphProperties?.NumberingProperties));
        Assert.True(conversion.Report.HasLoss);
        Assert.Contains(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfWordLegacyNumberingFlattened"
            && diagnostic.Action == RtfConversionAction.Flattened && diagnostic.Count == 1);
        Assert.Throws<RtfConversionLossException>(conversion.Report.RequireNoLoss);
        AssertArtifact(word, (string?)null);
        Assert.Equal(before, Native(source));
    }

    [Fact]
    public void Inherited_Legacy_Numbering_And_Cached_Discontinuities_Use_Destination_Instances() {
        RtfDocument source = RtfDocument.Create();
        source.AddStyle(10, "Legacy numbering").SetLegacyNumbering(numbering => {
            numbering.NumberStyle = RtfLegacyNumberingStyle.Decimal;
            numbering.StartAt = 5;
            numbering.TextAfter = ".";
        });
        source.AddParagraph("First").SetStyle(10).SetListText("5.\t");
        source.AddParagraph("Jumped").SetStyle(10).SetListText("9.\t");
        source.AddParagraph("Continued").SetStyle(10).SetListText("10.\t");
        string before = Native(source);

        RtfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        conversion.Report.RequireNoLoss();
        Assert.NotEqual(Id(word.Paragraphs[0]), Id(word.Paragraphs[1]));
        Assert.Equal(Id(word.Paragraphs[1]), Id(word.Paragraphs[2]));
        AssertArtifact(word, "5.", "9.", "10.");
        Assert.Equal(before, Native(source));
        Assert.All(source.Paragraphs, paragraph => {
            Assert.Null(paragraph.ListId);
            Assert.False(paragraph.LegacyNumbering.HasAnyValue);
        });
    }

    [Theory]
    [InlineData(RtfNoteKind.Footnote)]
    [InlineData(RtfNoteKind.Endnote)]
    [InlineData(RtfNoteKind.Annotation)]
    public void Note_Numbering_Continues_Within_Each_Note_And_Starts_Independently(RtfNoteKind kind) {
        RtfDocument source = RtfDocument.Create();
        source.AddStyle(10, "Authored note paragraph");
        RtfParagraph body = source.AddParagraph("Body");
        Legacy(body, 3);
        for (int index = 0; index < 2; index++) {
            var note = new RtfNote(kind);
            RtfParagraph first = note.AddParagraph("Note first");
            Legacy(first, 1);
            if (index == 0) {
                first.SetStyle(10);
                first.DirectKeepWithNext = true;
                first.SpaceAfterTwips = 240;
            }
            Legacy(note.AddParagraph("Note second"), 1);
            body.AddText("Reference").SetNote(note);
        }
        string before = Native(source);

        RtfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        conversion.Report.RequireNoLoss();
        AssertNoteBindings(word, kind);
        Assert.Empty(new OpenXmlValidator().Validate(word._wordprocessingDocument));
        using var bytes = new MemoryStream();
        word.Save(bytes);
        bytes.Position = 0;
        using WordDocument reopened = WordDocument.Load(bytes);
        AssertNoteBindings(reopened, kind);
        Assert.Empty(new OpenXmlValidator().Validate(reopened._wordprocessingDocument));
        Assert.Equal(before, Native(source));
    }

    private static void AssertNoteBindings(WordDocument word, RtfNoteKind kind) {
        var main = word._wordprocessingDocument.MainDocumentPart!;
        OpenXmlElement root = kind == RtfNoteKind.Footnote ? main.FootnotesPart!.Footnotes!
            : kind == RtfNoteKind.Endnote ? main.EndnotesPart!.Endnotes! : main.WordprocessingCommentsPart!.Comments!;
        Paragraph[] paragraphs = root.Descendants<Paragraph>().Where(paragraph => paragraph.InnerText.StartsWith("Note", StringComparison.Ordinal)).ToArray();
        Assert.Equal(4, paragraphs.Length);
        int[] ids = paragraphs.Select(paragraph => paragraph.ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value).ToArray();
        Assert.Equal(ids[0], ids[1]);
        Assert.Equal(ids[2], ids[3]);
        Assert.NotEqual(ids[0], ids[2]);
        Assert.NotEqual(Id(word.Paragraphs[0]), ids[0]);
        Assert.All(ids, id => Assert.Equal(1, LevelFor(word, id).StartNumberingValue!.Val!.Value));
        Assert.Equal("RtfP10", paragraphs[0].ParagraphProperties!.ParagraphStyleId!.Val!.Value);
        Assert.True(paragraphs[0].ParagraphProperties!.KeepNext!.Val!.Value);
        Assert.Equal("240", paragraphs[0].ParagraphProperties!.SpacingBetweenLines!.After!.Value);
        if (kind != RtfNoteKind.Annotation) {
            Assert.Equal(kind == RtfNoteKind.Footnote ? "FootnoteText" : "EndnoteText", paragraphs[2].ParagraphProperties!.ParagraphStyleId!.Val!.Value);
            Assert.Equal(2, kind == RtfNoteKind.Footnote ? root.Descendants<FootnoteReferenceMark>().Count() : root.Descendants<EndnoteReferenceMark>().Count());
        }
    }

    [Theory]
    [InlineData(RtfNoteKind.Footnote, "field-only")]
    [InlineData(RtfNoteKind.Footnote, "hyperlink-only")]
    [InlineData(RtfNoteKind.Footnote, "hyperlink-before-text")]
    [InlineData(RtfNoteKind.Endnote, "field-only")]
    [InlineData(RtfNoteKind.Endnote, "hyperlink-only")]
    [InlineData(RtfNoteKind.Endnote, "hyperlink-before-text")]
    public void Note_Reference_Precedes_Field_And_Hyperlink_Content(RtfNoteKind kind, string shape) {
        RtfDocument source = RtfDocument.Create();
        RtfParagraph body = source.AddParagraph("Body");
        body.AddBookmarkStart("LocalAnchor");
        body.AddBookmarkEnd("LocalAnchor");
        var note = new RtfNote(kind);
        RtfParagraph paragraph = note.AddParagraph();
        Legacy(paragraph, 1);
        paragraph.AddField(shape == "field-only" ? " PAGE " : @"HYPERLINK \l ""LocalAnchor""").AddText("Inline content");
        if (shape == "hyperlink-before-text") paragraph.AddText(" tail");
        body.Runs[0].SetNote(note);
        string before = Native(source);

        RtfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        conversion.Report.RequireNoLoss();
        AssertRtfInlineNote(word, kind, shape);
        Assert.Empty(new OpenXmlValidator().Validate(word._wordprocessingDocument));
        using WordDocument reopened = WordDocument.Load(new MemoryStream(word.ToBytes()));
        AssertRtfInlineNote(reopened, kind, shape);
        Assert.Empty(new OpenXmlValidator().Validate(reopened._wordprocessingDocument));
        Assert.Equal(before, Native(source));
    }

    private static void AssertRtfInlineNote(WordDocument word, RtfNoteKind kind, string shape) {
        var main = word._wordprocessingDocument.MainDocumentPart!;
        OpenXmlElement root = kind == RtfNoteKind.Footnote ? main.FootnotesPart!.Footnotes! : main.EndnotesPart!.Endnotes!;
        Paragraph paragraph = Assert.Single(root.Descendants<Paragraph>(), item => item.InnerText.StartsWith("Inline content", StringComparison.Ordinal));
        Assert.IsType<ParagraphProperties>(paragraph.ChildElements[0]);
        Assert.NotNull(paragraph.ParagraphProperties!.NumberingProperties);
        Run reference = Assert.IsType<Run>(paragraph.ChildElements[1]);
        if (kind == RtfNoteKind.Footnote) Assert.Single(reference.Elements<FootnoteReferenceMark>());
        else Assert.Single(reference.Elements<EndnoteReferenceMark>());
        if (shape == "field-only") {
            SimpleField field = Assert.IsType<SimpleField>(paragraph.ChildElements[2]);
            Assert.Equal(" PAGE ", field.Instruction!.Value);
        } else {
            Hyperlink hyperlink = Assert.IsType<Hyperlink>(paragraph.ChildElements[2]);
            Assert.Equal("LocalAnchor", hyperlink.Anchor!.Value);
        }
        Assert.Equal(shape == "hyperlink-before-text" ? "Inline content tail" : "Inline content", paragraph.InnerText);
    }

    private static void Legacy(RtfParagraph paragraph, int start) => paragraph.SetLegacyNumbering(numbering => {
        numbering.LevelKind = RtfLegacyNumberingLevelKind.Body;
        numbering.NumberStyle = RtfLegacyNumberingStyle.Decimal;
        numbering.StartAt = start;
        numbering.TextAfter = ".";
    });

    private static string Native(RtfDocument document) => document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });

    private static int Id(WordParagraph paragraph) {
        NumberingProperties? numbering = paragraph._paragraphProperties?.NumberingProperties;
        Assert.NotNull(numbering);
        Assert.NotNull(numbering.NumberingId?.Val);
        return numbering.NumberingId!.Val!.Value;
    }

    private static Level LevelFor(WordDocument word, int id) {
        Numbering numbering = word._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        NumberingInstance? instance = numbering.Elements<NumberingInstance>().SingleOrDefault(item => item.NumberID!.Value == id);
        Assert.True(instance != null, "Missing numbering instance " + id + ": " + numbering.OuterXml);
        int definition = instance!.AbstractNumId!.Val!.Value;
        return Assert.Single(numbering.Elements<AbstractNum>().Single(item => item.AbstractNumberId!.Value == definition).Elements<Level>());
    }

    private static void AssertArtifact(WordDocument word, params string?[] markers) {
        Assert.Empty(new OpenXmlValidator().Validate(word._wordprocessingDocument));
        using var bytes = new MemoryStream();
        word.Save(bytes);
        bytes.Position = 0;
        using WordDocument reopened = WordDocument.Load(bytes);
        Assert.Empty(new OpenXmlValidator().Validate(reopened._wordprocessingDocument));
        Assert.Equal(markers, reopened.Inspect().Sections.SelectMany(section => section.Elements)
            .OfType<WordParagraphInfo>().Select(paragraph => paragraph.ListMarker));
    }
}
