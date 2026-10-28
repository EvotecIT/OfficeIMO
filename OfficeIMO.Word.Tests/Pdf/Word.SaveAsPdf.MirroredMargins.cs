using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false, 1, WordFileFormat.Docx)]
    [InlineData(false, false, 2, WordFileFormat.Docx)]
    [InlineData(false, true, 1, WordFileFormat.Docx)]
    [InlineData(false, true, 2, WordFileFormat.Docx)]
    [InlineData(false, false, 1, WordFileFormat.Doc)]
    [InlineData(false, false, 2, WordFileFormat.Doc)]
    [InlineData(false, true, 1, WordFileFormat.Doc)]
    [InlineData(false, true, 2, WordFileFormat.Doc)]
    [InlineData(true, false, 1, WordFileFormat.Docx)]
    [InlineData(true, false, 1, WordFileFormat.Doc)]
    public void SaveAsPdf_MirroredMarginsFollowVisibleNumberingAndGutterDirection(
        bool atTop, bool onRight, int startNumber, WordFileFormat format) {
        using WordDocument document = WordDocument.Create();
        var section = document.Sections[0];
        section.PageSettings.PageSize = WordPageSize.Letter;
        section.Margins.Left = global::OfficeIMO.OfficeLength.Points(40D); section.Margins.Right = global::OfficeIMO.OfficeLength.Points(70D);
        section.Margins.Gutter = global::OfficeIMO.OfficeLength.Points(20D);
        section.RtlGutter = onRight;
        section.AddPageNumbering(startNumber);
        document.Settings.MirrorMargins = true;
        document.Settings.GutterAtTop = atTop;
        for (int index = 1; index <= 3; index++) {
            var paragraph = document.AddParagraph("MirrorMarker" + index);
            paragraph.FontFamily = "Arial"; paragraph.FontSize = 12;
            paragraph.PageBreakBeforeOverride = index > 1;
        }
        using var source = new MemoryStream(document.ToBytes(format));
        using WordDocument loaded = WordDocument.Load(source);
        using PdfPigDocument pdf = PdfPigDocument.Open(loaded.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(3, pdf.NumberOfPages);
        for (int index = 1; index <= 3; index++) {
            bool even = (startNumber + index - 1) % 2 == 0;
            // Word's producer output suppresses horizontal mirroring with a top gutter.
            double expectedLeft = atTop ? 40 : even ? (onRight ? 90 : 70) : (onRight ? 40 : 60);
            Assert.Equal(expectedLeft, FindWordStartX(pdf.GetPage(index), "MirrorMarker" + index), 3);
        }
    }
}
