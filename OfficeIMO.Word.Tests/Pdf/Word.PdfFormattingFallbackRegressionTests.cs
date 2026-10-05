using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public sealed class WordPdfFormattingFallbackRegressionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HiddenRunsCannotInflateTheRenderedBlankParagraph(bool minimum) {
        double Baseline(bool hidden) {
            using var document = WordDocument.Create();
            var blank = document.AddParagraph();
            blank._paragraph.ParagraphProperties = new W.ParagraphProperties(
                new W.ParagraphMarkRunProperties(new W.FontSize { Val = "22" }),
                new W.SpacingBetweenLines { Before = "0", After = "0", Line = minimum ? "400" : "240",
                    LineRule = minimum ? W.LineSpacingRuleValues.AtLeast : W.LineSpacingRuleValues.Auto });
            if (hidden) blank._paragraph.Append(new W.Run(new W.RunProperties(new W.Vanish(), new W.FontSize { Val = "144" }), new W.Text("Invisible")));
            document.AddParagraph("Marker");
            using var pdf = PdfPigDocument.Open(document.ToPdfBytes(Options()));
            Assert.DoesNotContain("Invisible", pdf.GetPage(1).Text);
            return pdf.GetPage(1).Letters.First(letter => letter.Value == "M").StartBaseLine.Y;
        }
        Assert.Equal(Baseline(false), Baseline(true), precision: 3);
    }

    [Theory]
    [InlineData(false, false, 14)]
    [InlineData(false, true, 16)]
    [InlineData(true, false, 14)]
    [InlineData(false, false, 0)]
    public void GeneratedNoteMarkerUsesItsSourceReferenceSizeInAMixedParagraph(bool table, bool endnote, int referencePoints) {
        using var document = WordDocument.Create();
        var paragraph = table ? document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0] : document.AddParagraph();
        paragraph._paragraph.RemoveAllChildren<W.Run>();
        paragraph._paragraph.Append(Text("Small ", 8), Text("Body", 11));
        if (endnote) paragraph.AddEndNote("Endnote text");
        else paragraph.AddFootNote("Footnote text");
        OpenXmlElement reference = endnote ? paragraph._paragraph.Descendants<W.EndnoteReference>().Single() :
            paragraph._paragraph.Descendants<W.FootnoteReference>().Single();
        var referenceRun = (W.Run)reference.Parent!;
        referenceRun.RunProperties ??= new W.RunProperties();
        referenceRun.RunProperties.FontSize = referencePoints > 0 ? new W.FontSize { Val = (referencePoints * 2).ToString() } : null;
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(Options()));
        double baseline = pdf.GetPage(1).Letters.First(letter => letter.Value == "B").StartBaseLine.Y;
        var marker = Assert.Single(pdf.GetPage(1).Letters, letter => letter.Value == "1" && letter.StartBaseLine.Y > baseline + .1);
        Assert.Equal((referencePoints > 0 ? referencePoints : 11D) * .65D, marker.PointSize, precision: 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ConfiguredDefaultTableFontAndLeadingReachTheRenderedUnstyledCell(bool authored) {
        using var document = WordDocument.Create();
        var styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle?.RemoveAllChildren<W.SpacingBetweenLines>();
        foreach (var style in styles.Elements<W.Style>().Where(style => style.StyleId?.Value == "Normal"))
            style.StyleParagraphProperties?.RemoveAllChildren<W.SpacingBetweenLines>();
        var table = document.AddTable(1, 1);
        table._tableProperties!.TableStyle?.Remove();
        var paragraph = table.Rows[0].Cells[0].Paragraphs[0];
        paragraph._paragraph.RemoveAllChildren<W.Run>();
        paragraph._paragraph.ParagraphProperties = new W.ParagraphProperties(new W.SpacingBetweenLines { Before = "0", After = "0" });
        paragraph._paragraph.Append(new W.Run(new W.Text("A"), new W.Break(), new W.Text("B"), new W.Break(), new W.Text("C")));
        if (authored) {
            paragraph._paragraph.Elements<W.Run>().Single().RunProperties = new W.RunProperties(new W.FontSize { Val = "36" });
            paragraph._paragraph.ParagraphProperties.SpacingBetweenLines!.Line = "480";
            paragraph._paragraph.ParagraphProperties.SpacingBetweenLines.LineRule = W.LineSpacingRuleValues.Exact;
        }
        var options = Options();
        options.PdfOptions!.DefaultTableStyle = new PdfTableStyle {
            FontSize = 12.5, LineHeight = 1.4, HeaderRowCount = 0, CellPaddingX = 0, CellPaddingY = 0,
            BorderWidth = 0, SpacingBefore = 0, SpacingAfter = 0
        };
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(options));
        var letters = pdf.GetPage(1).Letters.Where(letter => letter.Value is "A" or "B" or "C").ToArray();
        Assert.Equal(3, letters.Length);
        Assert.All(letters, letter => Assert.Equal(authored ? 18D : 12.5D, letter.PointSize, precision: 3));
        Assert.Equal(authored ? 24D : 17.5D, letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y, precision: 3);
        Assert.Equal(authored ? 24D : 17.5D, letters[1].StartBaseLine.Y - letters[2].StartBaseLine.Y, precision: 3);
    }

    [Fact]
    public void VmlFormattedBreakKeepsItsOwnHeightWithoutChangingTheOtherLines() {
        using var document = WordDocument.Create();
        var content = new W.Paragraph(Text("A", 11),
            new W.Run(new W.RunProperties(new W.FontSize { Val = "64" }), new W.Break()),
            Text("B", 11), new W.Run(new W.RunProperties(new W.FontSize { Val = "22" }), new W.Break()), Text("C", 11));
        var shape = new V.Shape(new V.TextBox(new W.TextBoxContent(content))) {
            Id = "FormattedBreakBox", Type = "#_x0000_t202", Style = "position:absolute;left:72pt;top:72pt;width:420pt;height:240pt"
        };
        var cover = new W.SdtBlock(new W.SdtProperties(new W.SdtContentDocPartObject(
            new W.DocPartGallery { Val = "Cover Pages" }, new W.DocPartUnique())),
            new W.SdtContentBlock(new W.Paragraph(new W.Run(new W.Picture(shape)))));
        var body = document._wordprocessingDocument.MainDocumentPart!.Document.Body!;
        body.InsertBefore(cover, body.Elements<W.SectionProperties>().FirstOrDefault());
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(Options()));
        var letters = pdf.GetPage(1).Letters.Where(letter => letter.Value is "A" or "B" or "C").ToArray();
        Assert.Equal(3, letters.Length);
        Assert.Equal(13.2D, letters[1].StartBaseLine.Y - letters[2].StartBaseLine.Y, precision: 3);
        // The formatted break enlarges its own line box; font ascent also
        // changes that line's baseline, so the baseline delta is not 32 * 1.2.
        Assert.InRange(letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y, 13.3D, 38.5D);
    }

    private static W.Run Text(string text, int points) => new(new W.RunProperties(new W.FontSize { Val = (points * 2).ToString() }),
        new W.Text(text) { Space = SpaceProcessingModeValues.Preserve });
    private static WordToPdfOptions Options() => new() {
        IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
        PdfOptions = new PdfOptions { CompressContentStreams = false }
    };
}
