using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(1, 12, false)]
    [InlineData(2, 12, false)]
    [InlineData(3, 12, false)]
    [InlineData(1, 24, false)]
    [InlineData(2, 24, false)]
    [InlineData(3, 24, false)]
    [InlineData(1, 12, true)]
    [InlineData(2, 12, true)]
    [InlineData(3, 12, true)]
    [InlineData(1, 24, true)]
    [InlineData(2, 24, true)]
    [InlineData(3, 24, true)]
    public void SaveAsPdf_EmptySectionBookmarkPrecedesItsParagraphLine(int columns, int size, bool nextPage) {
        using WordDocument source = WordDocument.Create();
        source.Settings.FontFamily = "Arial";
        source.Settings.FontSize = 12;
        W.Body body = source._wordprocessingDocument.MainDocumentPart!.Document.Body!;
        W.SectionProperties final = body.Elements<W.SectionProperties>().Single();
        var first = (W.SectionProperties)final.CloneNode(true);
        first.AddChild(new W.Columns { ColumnCount = (short)columns, Space = "720", Separator = true }, true);
        var mark = new W.Paragraph(
            new W.ParagraphProperties(new W.ParagraphMarkRunProperties(new W.FontSize { Val = (size * 2).ToString() }), first),
            new W.BookmarkStart { Id = "0", Name = "EmptySectionTarget" }, new W.BookmarkEnd { Id = "0" },
            new W.BookmarkStart { Id = "1", Name = "SecondEmptyTarget" }, new W.BookmarkEnd { Id = "1" });
        body.InsertBefore(mark, final);
        final.AddChild(new W.SectionType { Val = nextPage ? W.SectionMarkValues.NextPage : W.SectionMarkValues.Continuous }, true);
        source.AddParagraph("BODY");
        source.AddHyperLink("Go to target", "EmptySectionTarget");
        source.AddHyperLink("Go to second target", "SecondEmptyTarget");
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        string xml = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using PdfDocument model = document.ToPdf(new WordPdfExportOptions { ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() }, out OfficeConversionReport report);
        Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Code.StartsWith("native-bookmark-target", StringComparison.Ordinal));
        byte[] bytes = model.ToBytes(new PdfSaveOptions { AllowLoss = true });
        using var pdf = PdfPigDocument.Open(bytes);
        var logical = PdfDocumentReadResult.Load(bytes);
        var destination = Assert.Single(logical.NamedDestinations, d => d.Name == "EmptySectionTarget");
        Assert.Equal(1, destination.PageNumber);
        // Native Word targets the empty paragraph's start, before its line is consumed.
        Assert.Equal(pdf.GetPage(1).Height - 72D, destination.DestinationTop!.Value, 3);
        Assert.Single(logical.GetLinksByDestinationName("EmptySectionTarget"));
        var second = Assert.Single(logical.NamedDestinations, d => d.Name == "SecondEmptyTarget");
        Assert.Single(logical.GetLinksByDestinationName("SecondEmptyTarget"));
        Assert.Equal(destination.PageNumber, second.PageNumber);
        Assert.Equal(destination.DestinationTop, second.DestinationTop);
        Assert.Equal(xml, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }

    [Theory]
    [InlineData(false, "pure")]
    [InlineData(false, "styled")]
    [InlineData(false, "hidden")]
    [InlineData(false, "mark-size")]
    [InlineData(true, "pure")]
    [InlineData(true, "styled")]
    public void SaveAsPdf_EmptySectionRetainsItsOwnParagraphLine(bool precedingText, string mode) {
        using WordDocument source = CreateDocument();
        if (precedingText) AddLine(source, "P");
        W.Body body = source._wordprocessingDocument.MainDocumentPart!.Document.Body!;
        W.SectionProperties final = body.Elements<W.SectionProperties>().Single();
        var properties = new W.ParagraphProperties((W.SectionProperties)final.CloneNode(true));
        ConfigureMark(properties);
        body.InsertBefore(new W.Paragraph(properties), final);
        final.RemoveAllChildren<W.SectionType>();
        final.AddChild(new W.SectionType { Val = W.SectionMarkValues.Continuous }, true);
        AddLine(source, "A").AddBreak(WordBreakType.Page);
        AddLine(source, "B");
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        string before = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;

        using WordDocument reference = CreateDocument();
        if (precedingText) AddLine(reference, "P");
        // Native Word gives an otherwise empty section its paragraph line.
        // A terminator after real text adds no separate line.
        else if (mode != "hidden") {
            WordParagraph blank = reference.AddParagraph();
            ConfigureMark(blank._paragraph.ParagraphProperties ??= new W.ParagraphProperties());
        }
        AddLine(reference, "A").AddBreak(WordBreakType.Page);
        AddLine(reference, "B");
        var options = new WordPdfExportOptions { ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        using var actual = PdfPigDocument.Open(document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        using var expected = PdfPigDocument.Open(reference.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(2, actual.NumberOfPages);
        for (int page = 1; page <= 2; page++) Assert.Equal(
            expected.GetPage(page).Letters.Select(letter => (letter.Value, letter.StartBaseLine)),
            actual.GetPage(page).Letters.Select(letter => (letter.Value, letter.StartBaseLine)));
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);

        void ConfigureMark(W.ParagraphProperties properties) {
            if (mode == "styled") properties.ParagraphStyleId = new W.ParagraphStyleId { Val = "Normal" };
            if (mode == "hidden") properties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties(new W.Vanish());
            if (mode == "mark-size") properties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties(new W.FontSize { Val = "48" });
        }
        static WordDocument CreateDocument() {
            var result = WordDocument.Create();
            result.CompatibilitySettings.CompatibilityMode = WordCompatibilityMode.Word2013;
            result.Settings.FontFamily = "Arial";
            result.Settings.FontSize = 12;
            return result;
        }
        static WordParagraph AddLine(WordDocument target, string text) {
            WordParagraph paragraph = target.AddParagraph(text);
            paragraph.LineSpacingPoints = 12;
            paragraph.LineSpacingRule = WordLineSpacingRule.Exact;
            paragraph.LineSpacingBeforePoints = paragraph.LineSpacingAfterPoints = 0;
            return paragraph;
        }
    }
}
