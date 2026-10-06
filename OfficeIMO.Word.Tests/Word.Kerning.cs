using OfficeIMO.Word;
using Xunit;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void NativeDocKerningAndFontSizeProduceValidStyleAndParagraphMarkProperties(bool paragraphMark, bool additionalFormatting) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph("AV formatted text");
        var formatting = new StyleRunProperties();
        formatting.AddChild(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, true);
        formatting.AddChild(new Kern { Val = 16U }, true);
        formatting.AddChild(new FontSize { Val = "48" }, true);
        if (additionalFormatting) {
            formatting.AddChild(new Bold(), true);
            formatting.AddChild(new Caps(), true);
            formatting.AddChild(new Spacing { Val = 20 }, true);
            formatting.AddChild(new Underline { Val = UnderlineValues.Single }, true);
            formatting.AddChild(new Languages { Val = "en-US" }, true);
        }

        if (paragraphMark) {
            var mark = new ParagraphMarkRunProperties();
            foreach (OpenXmlElement property in formatting.ChildElements) {
                mark.AddChild(property.CloneNode(true), true);
            }
            var paragraphProperties = paragraph._paragraphProperties
                ?? paragraph._paragraph.PrependChild(new ParagraphProperties());
            paragraphProperties.AddChild(mark, true);
        } else {
            const string styleId = "KerningSchemaOrder";
            var style = new Style {
                Type = StyleValues.Paragraph, StyleId = styleId, CustomStyle = true
            };
            style.Append(new StyleName { Val = "Kerning Schema Order" });
            style.StyleRunProperties = formatting;
            source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);
            paragraph.SetStyleId(styleId);
        }

        using WordDocument reloaded = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        OpenXmlCompositeElement properties = paragraphMark
            ? reloaded.Paragraphs[0]._paragraphProperties!.GetFirstChild<ParagraphMarkRunProperties>()!
            : reloaded._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<Style>()
                .Single(style => style.StyleId == reloaded.Paragraphs[0].StyleId).StyleRunProperties!;
        Assert.Equal(16U, properties.GetFirstChild<Kern>()!.Val!.Value);
        Assert.Equal("48", properties.GetFirstChild<FontSize>()!.Val!.Value);
        if (additionalFormatting) {
            Assert.Equal("en-US", properties.GetFirstChild<Languages>()!.Val!.Value);
            Assert.Equal(20, properties.GetFirstChild<Spacing>()!.Val!.Value);
        }
        var validator = new OpenXmlValidator();
        Assert.Empty(validator.Validate(properties));
        using WordDocument docx = WordDocument.Load(new MemoryStream(reloaded.ToBytes()));
        Assert.Empty(docx.ValidateDocument());
    }

    [Theory]
    [InlineData(false, 0D)]
    [InlineData(false, 1.5D)]
    [InlineData(false, 1638D)]
    [InlineData(true, 0D)]
    [InlineData(true, 1.5D)]
    [InlineData(true, 1638D)]
    public void RunKerningThresholdPreservesNativeHalfPoints(bool nativeDoc, double threshold) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("AB").KerningMinimumFontSizePoints = threshold;
        using WordDocument reloaded = WordDocument.Load(new MemoryStream(nativeDoc
            ? source.ToBytes(WordFileFormat.Doc) : source.ToBytes()));
        Assert.Equal(threshold, reloaded.Paragraphs[0].KerningMinimumFontSizePoints);
    }

    [Fact]
    public void RunKerningThresholdCanReturnToInheritanceAndRejectsInvalidValuesAtomically() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("AB");
        paragraph.KerningMinimumFontSizePoints = 1.25D;
        Assert.Equal(1.5D, paragraph.KerningMinimumFontSizePoints);
        foreach (double invalid in new[] { -0.5D, 1638.5D, double.NaN, double.PositiveInfinity }) {
            Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.KerningMinimumFontSizePoints = invalid);
            Assert.Equal(1.5D, paragraph.KerningMinimumFontSizePoints);
        }
        paragraph.KerningMinimumFontSizePoints = null;
        using WordDocument reloaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Null(reloaded.Paragraphs[0].KerningMinimumFontSizePoints);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RunKerningThresholdTargetsTheHyperlinkWithoutChangingAdjacentText(bool nativeDoc) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Prefix ");
        paragraph.KerningMinimumFontSizePoints = 8D;
        WordParagraph link = paragraph.AddHyperLink("AV", new Uri("https://example.com/kerning"));
        link.KerningMinimumFontSizePoints = 1.5D;
        Assert.Equal(1.5D, link.KerningMinimumFontSizePoints);
        Assert.Equal(16U, paragraph._paragraph.Elements<Run>()
            .First().RunProperties!.Kern!.Val!.Value);

        using WordDocument reloaded = WordDocument.Load(new MemoryStream(nativeDoc
            ? document.ToBytes(WordFileFormat.Doc) : document.ToBytes()));
        WordParagraph reloadedLink = Assert.Single(reloaded.Paragraphs, item => item.IsHyperLink);
        Assert.Equal(1.5D, reloadedLink.KerningMinimumFontSizePoints);
        Assert.Equal(new Uri("https://example.com/kerning"), reloadedLink.Hyperlink!.Uri);
        Assert.Equal(8D, reloaded.Paragraphs.First(item => item.Text == "Prefix ").KerningMinimumFontSizePoints);
        reloadedLink.KerningMinimumFontSizePoints = null;
        Assert.Null(reloadedLink.KerningMinimumFontSizePoints);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RunKerningThresholdPreservesTablesAndSeparateStories(bool nativeDoc) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body");
        document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].AddText("Cell")
            .KerningMinimumFontSizePoints = 0D;
        document.HeaderDefaultOrCreate.AddParagraph("Header").KerningMinimumFontSizePoints = 1.5D;
        document.FooterDefaultOrCreate.AddParagraph("Footer").KerningMinimumFontSizePoints = 13D;
        document.AddParagraph("Reference").AddFootNote("Footnote").FootNote!.Paragraphs!
            .Single(item => item.Text == "Footnote").KerningMinimumFontSizePoints = 7.5D;
        document.AddParagraph("Reference").AddEndNote("Endnote").EndNote!.Paragraphs!
            .Single(item => item.Text == "Endnote").KerningMinimumFontSizePoints = 9D;

        using WordDocument reloaded = WordDocument.Load(new MemoryStream(nativeDoc
            ? document.ToBytes(WordFileFormat.Doc) : document.ToBytes()));
        Assert.Equal(0D, reloaded.Tables[0].Rows[0].Cells[0].Paragraphs.Single(item => item.Text == "Cell")
            .KerningMinimumFontSizePoints);
        Assert.Equal(1.5D, reloaded.Sections[0].Header.Default!.Paragraphs.Single(item => item.Text == "Header")
            .KerningMinimumFontSizePoints);
        Assert.Equal(13D, reloaded.Sections[0].Footer.Default!.Paragraphs.Single(item => item.Text == "Footer")
            .KerningMinimumFontSizePoints);
        Assert.Equal(7.5D, Assert.Single(reloaded.FootNotes).Paragraphs!.Single(item => item.Text == "Footnote")
            .KerningMinimumFontSizePoints);
        Assert.Equal(9D, Assert.Single(reloaded.EndNotes).Paragraphs!.Single(item => item.Text == "Endnote")
            .KerningMinimumFontSizePoints);
    }
}
