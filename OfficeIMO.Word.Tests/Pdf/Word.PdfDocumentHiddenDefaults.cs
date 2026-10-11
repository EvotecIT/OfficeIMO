using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("body", "none")]
    [InlineData("body", "character")]
    [InlineData("table", "none")]
    [InlineData("table", "paragraph")]
    [InlineData("header", "none")]
    [InlineData("header", "direct")]
    [InlineData("footer", "none")]
    [InlineData("footer", "character")]
    public void SaveAsPdf_DocumentHiddenDefaultControlsTextVisibility(string location, string visibleOverride) {
        using WordDocument document = CreateDocumentHiddenDefaultFixture();
        W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new W.Style(new W.StyleName { Val = "Visible Override" },
            new W.BasedOn { Val = "Normal" }, new W.StyleRunProperties(new W.Vanish { Val = false })) {
            StyleId = "VisibleOverride", Type = visibleOverride == "paragraph" ? W.StyleValues.Paragraph : W.StyleValues.Character
        });
        WordParagraph paragraph = AddDocumentHiddenDefaultParagraph(document, location);
        if (visibleOverride == "paragraph") paragraph.SetStyleId("VisibleOverride");
        paragraph._paragraph.Append(new W.Run(new W.Text("SECRET")));
        var overrideProperties = new W.RunProperties();
        if (visibleOverride == "character") overrideProperties.AddChild(new W.RunStyle { Val = "VisibleOverride" }, true);
        if (visibleOverride == "direct") overrideProperties.AddChild(new W.Vanish { Val = false }, true);
        paragraph._paragraph.Append(new W.Run(overrideProperties, new W.Text("OVERRIDE")),
            new W.Run(new W.RunProperties(new W.Vanish { Val = false }), new W.Text("VISIBLE")));
        HideJoinMark(paragraph, false);
        Assert.Empty(document.ValidateDocument());
        using var pdf = OpenJoinedParagraphPdf(document);
        string text = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        Assert.Contains("VISIBLE", text);
        Assert.Equal(visibleOverride == "paragraph", text.Contains("SECRET"));
        Assert.Equal(visibleOverride != "none", text.Contains("OVERRIDE"));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SaveAsPdf_DocumentHiddenDefaultControlsBlankMarkHeight(bool table, bool visibleMark) {
        double baseline = Gap(includeBlank: false);
        double actual = Gap(includeBlank: true);
        Assert.Equal(baseline * (visibleMark ? 2D : 1D), actual, 3);

        double Gap(bool includeBlank) {
            using WordDocument document = CreateDocumentHiddenDefaultFixture();
            WordTableCell? cell = table ? document.AddTable(1, 1).Rows[0].Cells[0] : null;
            if (cell != null) cell.Paragraphs[0].Remove();
            WordParagraph Add(string text) => cell == null ? document.AddParagraph(text) : cell.AddParagraph(text);
            WordParagraph alpha = Add("A"); alpha.Hidden = false; HideJoinMark(alpha, false);
            if (includeBlank) {
                WordParagraph blank = Add(string.Empty);
                blank._paragraph.RemoveAllChildren<W.Run>();
                if (visibleMark) HideJoinMark(blank, false);
            }
            WordParagraph beta = Add("B"); beta.Hidden = false; HideJoinMark(beta, false);
            Assert.Empty(document.ValidateDocument());
            using var pdf = OpenJoinedParagraphPdf(document);
            var letters = pdf.GetPage(1).Letters;
            return Assert.Single(letters, letter => letter.Value == "A").StartBaseLine.Y -
                Assert.Single(letters, letter => letter.Value == "B").StartBaseLine.Y;
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SaveAsPdf_DocumentHiddenDefaultControlsParagraphJoins(bool table, bool visibleMark) {
        using WordDocument document = CreateDocumentHiddenDefaultFixture();
        WordTableCell? cell = table ? document.AddTable(1, 1).Rows[0].Cells[0] : null;
        if (cell != null) cell.Paragraphs[0].Remove();
        WordParagraph Add(string text) => cell == null ? document.AddParagraph(text) : cell.AddParagraph(text);
        WordParagraph alpha = Add("ALPHA"); alpha.Hidden = false;
        if (visibleMark) HideJoinMark(alpha, false);
        WordParagraph beta = Add("BETA"); beta.Hidden = false; HideJoinMark(beta, false);
        string source = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        Assert.Empty(document.ValidateDocument());
        using var pdf = OpenJoinedParagraphPdf(document);
        var letters = pdf.GetPage(1).Letters;
        double gap = letters.First(letter => letter.Value == "A").StartBaseLine.Y -
            Assert.Single(letters, letter => letter.Value == "B").StartBaseLine.Y;
        if (visibleMark) Assert.True(gap > 0D);
        else Assert.Equal(0D, gap, 3);
        Assert.Equal(source, document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml);
    }

    private static WordDocument CreateDocumentHiddenDefaultFixture() {
        WordDocument document = CreateJoinedParagraphDocument();
        document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!.AddChild(new W.Vanish(), true);
        return document;
    }

    private static WordParagraph AddDocumentHiddenDefaultParagraph(WordDocument document, string location) {
        WordParagraph paragraph = location switch {
            "header" => document.HeaderDefaultOrCreate.AddParagraph(),
            "footer" => document.FooterDefaultOrCreate.AddParagraph(),
            "table" => document.AddTable(1, 1).Rows[0].Cells[0].AddParagraph(removeExistingParagraphs: true),
            _ => document.AddParagraph()
        };
        paragraph._paragraph.RemoveAllChildren<W.Run>();
        return paragraph;
    }
}
