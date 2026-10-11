using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using UglyToad.PdfPig.Content;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf {
    public sealed class PdfNaturalFontLineAlignmentTests {
        [Theory]
        [InlineData(false, "Top", 7.08D)]
        [InlineData(true, "Top", 7.08D)]
        [InlineData(false, "Center", 2.64D)]
        [InlineData(true, "Center", 2.64D)]
        [InlineData(false, "Bottom", -.96D)]
        [InlineData(true, "Bottom", -.96D)]
        [InlineData(false, "Baseline", 0D)]
        [InlineData(true, "Baseline", 0D)]
        public void NaturalAndMinimumLinesAlignEachRunByItsOwnMetrics(bool minimum, string alignmentName, double difference) {
            PdfFontLineAlignment alignment = (PdfFontLineAlignment)Enum.Parse(typeof(PdfFontLineAlignment), alignmentName);
            PdfOptions options = new PdfOptions();
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Natural device metrics", ManagedTextShapingTestAssets.CreateFontWithDeviceMetrics()));
            PdfLineSpacing spacing = (minimum ? PdfLineSpacing.AtLeast(30, 1.2D) : PdfLineSpacing.Multiple(1.2D))
                .WithFontLineBoxBaseline(1.2D).WithFontAlignment(alignment, .12D);
            PdfDocument document = PdfDocument.Create(options);
            document.Paragraph(paragraph => paragraph.FontFamily("Natural device metrics").FontSize(8).Text("A")
                .FontSize(16).Text("B"), style: new PdfParagraphStyle { LineSpacing = spacing });
            using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes());
            IReadOnlyList<Letter> letters = pdf.GetPage(1).Letters;
            Letter first = Assert.Single(letters, letter => letter.Value == "A");
            Letter second = Assert.Single(letters, letter => letter.Value == "B");
            Assert.InRange(first.StartBaseLine.Y - second.StartBaseLine.Y, difference - .002D, difference + .002D);
        }
    }
}
