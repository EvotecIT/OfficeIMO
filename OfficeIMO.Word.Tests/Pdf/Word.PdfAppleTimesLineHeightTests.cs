using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void PdfExport_AppleTimesKeepsSingleLineSpacing() {
        const string appleTimes = "/System/Library/Fonts/Times.ttc";
        if (!File.Exists(appleTimes)) return;
        string output = Path.Combine(_directoryWithFiles, "AppleTimesLines.pdf");
        using (WordDocument document = WordDocument.Create()) {
            document.AddParagraph("First").SetFontFamily("Times").SetFontSize(OfficeLength.Points(12));
            document.AddParagraph("Second").SetFontFamily("Times").SetFontSize(OfficeLength.Points(12));
            document.SaveAsPdf(output);
        }

        using var pdf = PdfPigDocument.Open(output);
        var words = pdf.GetPage(1).GetWords().ToArray();
        double advance = words.Single(word => word.Text == "First").BoundingBox.Bottom - words.Single(word => word.Text == "Second").BoundingBox.Bottom;
        // The default style adds 1.08 line spacing and 8pt after: 13.8 x 1.08 + 8.
        Assert.InRange(advance, 22D, 24D);
    }
}
