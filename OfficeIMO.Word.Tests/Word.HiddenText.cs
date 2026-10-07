using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void LegacyDoc_AdjacentUnformattedRunsDoNotEraseExplicitOffToggleOverrides() {
        using WordDocument source = WordDocument.Create();
        W.OnOffType[] properties = {
            new W.Bold(), new W.Italic(), new W.Strike(), new W.DoubleStrike(),
            new W.Outline(), new W.Shadow(), new W.Emboss(), new W.Imprint(),
            new W.Vanish(), new W.NoProof(), new W.Caps(), new W.SmallCaps()
        };
        foreach (W.OnOffType property in properties) {
            string styleId = "Enabled" + property.LocalName;
            var style = new W.Style {
                StyleId = styleId, Type = W.StyleValues.Paragraph, CustomStyle = true,
                StyleName = new W.StyleName { Val = styleId }, BasedOn = new W.BasedOn { Val = "Normal" },
                StyleRunProperties = new W.StyleRunProperties(property.CloneNode(true))
            };
            source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);
            source.AddParagraph("Inherited" + property.LocalName).SetStyleId(styleId);
            WordParagraph off = source.AddParagraph("ExplicitOff" + property.LocalName).SetStyleId(styleId);
            var disabled = (W.OnOffType)property.CloneNode(true);
            disabled.Val = false;
            off._runProperties = new W.RunProperties(disabled);
        }
        Assert.Empty(source.ValidateDocument());

        using WordDocument reopened = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        var missing = new List<string>();
        foreach (W.OnOffType property in properties) {
            WordParagraph off = reopened.Paragraphs.Single(p => p.Text == "ExplicitOff" + property.LocalName);
            W.OnOffType? saved = off._runProperties?.Elements<W.OnOffType>().SingleOrDefault(p => p.GetType() == property.GetType());
            if (saved?.Val?.Value != false) missing.Add(property.LocalName);
        }
        Assert.Empty(missing);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData(WordFileFormat.Docx, true)]
    [InlineData(WordFileFormat.Docx, false)]
    [InlineData(WordFileFormat.Doc, true)]
    [InlineData(WordFileFormat.Doc, false)]
    public void HiddenText_ExplicitFormattingRoundTripsWithoutLosingText(WordFileFormat format, bool hidden) {
        using WordDocument source = WordDocument.Create();
        WordParagraph run = source.AddParagraph("Retained hidden content");
        run.FontSize = 12;
        run.Hidden = hidden;
        Assert.Equal(hidden, run.Hidden);
        Assert.Empty(source.ValidateDocument());

        using WordDocument reopened = WordDocument.Load(new MemoryStream(source.ToBytes(format)));
        WordParagraph saved = Assert.Single(reopened.Paragraphs);
        Assert.Equal("Retained hidden content", saved.Text);
        Assert.Equal(hidden, saved.Hidden);
        Assert.Equal(12, saved.FontSize);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void HiddenText_ImportedOnAndClearRetainOtherRunFormattingAndDoNotCreateEmptyRuns() {
        using WordDocument source = WordDocument.Create();
        WordParagraph empty = source.AddParagraph();
        string emptyXml = empty._paragraph.OuterXml;
        Assert.Null(empty.Hidden);
        empty.Hidden = null;
        Assert.Equal(emptyXml, empty._paragraph.OuterXml);

        WordParagraph run = source.AddParagraph("Hidden template text");
        run.Bold = true;
        run._runProperties!.Vanish = new W.Vanish();
        Assert.True(run.Hidden);
        run.Hidden = false;
        Assert.False(run.Hidden);
        run.Hidden = null;
        Assert.Null(run.Hidden);
        Assert.True(run.Bold);
        using WordDocument reopened = WordDocument.Load(new MemoryStream(source.ToBytes()));
        WordParagraph saved = reopened.Paragraphs.Single(p => p.Text == "Hidden template text");
        Assert.Null(saved.Hidden);
        Assert.True(saved.Bold);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void HiddenText_ContentControlInsideHyperlinkKeepsItsOwnFormatting() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph("Anchor");
        paragraph._run!.Remove();
        var hyperlink = new W.Hyperlink(
            new W.Run(new W.Text("Before ")),
            new W.SdtRun(new W.SdtProperties(), new W.SdtContentRun(new W.Run(new W.Text("Inner")))),
            new W.Run(new W.Text(" After"))) { Anchor = "target" };
        paragraph._paragraph.Append(hyperlink);

        WordParagraph[] runs = paragraph.GetRuns().ToArray();
        runs[0].Hidden = false;
        runs[1].Hidden = true;
        Assert.Null(runs[2].Hidden);
        Assert.False(hyperlink.Elements<W.Run>().First().RunProperties!.Vanish!.Val!.Value);
        Assert.True(hyperlink.Descendants<W.SdtRun>().Single().Descendants<W.Run>().Single().RunProperties!.Vanish!.Val!.Value);
        Assert.Null(hyperlink.Elements<W.Run>().Last().RunProperties);
        Assert.Equal("Before Inner After", hyperlink.InnerText);

        using WordDocument reopened = WordDocument.Load(new MemoryStream(source.ToBytes()));
        WordParagraph[] saved = reopened.Paragraphs.Single().GetRuns().ToArray();
        Assert.False(saved[0].Hidden);
        Assert.True(saved[1].Hidden);
        Assert.Null(saved[2].Hidden);
        saved[1].Hidden = null;
        Assert.False(saved[0].Hidden);
        Assert.Null(saved[1].Hidden);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData(WordFileFormat.Docx)]
    [InlineData(WordFileFormat.Doc)]
    public void HiddenText_ExplicitVisibleOverridesStyleAndClearRestoresPdfInheritance(WordFileFormat format) {
        using WordDocument source = WordDocument.Create();
        var style = new W.Style {
            StyleId = "HiddenTextStyle", Type = W.StyleValues.Paragraph, CustomStyle = true,
            StyleName = new W.StyleName { Val = "Hidden Text Style" },
            BasedOn = new W.BasedOn { Val = "Normal" },
            StyleRunProperties = new W.StyleRunProperties(new W.Vanish())
        };
        source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);
        source.AddParagraph("VISIBLECONTROL");
        WordParagraph inherited = source.AddParagraph("INHERITEDSECRET").SetStyleId("HiddenTextStyle");
        WordParagraph visible = source.AddParagraph("VISIBLEOVERRIDE").SetStyleId("HiddenTextStyle");
        visible.Hidden = false;
        WordParagraph cleared = source.AddParagraph("CLEAREDSECRET").SetStyleId("HiddenTextStyle");
        cleared.Hidden = false;
        cleared.Hidden = null;
        source.AddParagraph("DIRECTSECRET").Hidden = true;
        Assert.Null(inherited.Hidden);
        Assert.Null(cleared.Hidden);
        Assert.Empty(source.ValidateDocument());

        using WordDocument reopened = WordDocument.Load(new MemoryStream(source.ToBytes(format)));
        string text = OfficeIMO.Pdf.PdfTextExtractor.ExtractAllText(reopened.ToPdfBytes(
            new WordToPdfOptions { IncludePageNumbers = false, FontFamily = "Helvetica" }));
        Assert.Contains("VISIBLECONTROL", text);
        Assert.Contains("VISIBLEOVERRIDE", text);
        Assert.DoesNotContain("INHERITEDSECRET", text);
        Assert.DoesNotContain("CLEAREDSECRET", text);
        Assert.DoesNotContain("DIRECTSECRET", text);
    }
}
