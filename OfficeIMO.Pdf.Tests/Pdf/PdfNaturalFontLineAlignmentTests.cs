using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using UglyToad.PdfPig.Content;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf {
    public sealed class PdfNaturalFontLineAlignmentTests {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void AlignedNaturalLinesReserveTheirCombinedFontBox(bool minimum) {
            PdfOptions options = new PdfOptions();
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Natural line bounds", ManagedTextShapingTestAssets.CreateFontWithDeviceMetrics()));
            PdfLineSpacing spacing = (minimum ? PdfLineSpacing.AtLeast(1, 1D) : PdfLineSpacing.Multiple(1D))
                .WithFontLineBoxBaseline(1D).WithFontAlignment(PdfFontLineAlignment.Top, .12D);
            PdfDocument document = PdfDocument.Create(options);
            document.Compose(builder => builder.Content(content => content.Paragraph(paragraph => paragraph
                .FontFamily("Natural line bounds").FontSize(8).Text("A").FontSize(16).Text("B")
                .Text("\n").FontSize(8).Text("A").FontSize(16).Text("B"),
                style: new PdfParagraphStyle { LineSpacing = spacing })));
            using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes());
            Letter[] letters = pdf.GetPage(1).Letters.Where(letter => letter.Value == "A").ToArray();
            Assert.Equal(2, letters.Length);
            // These device bounds exceed the font-unit ascent. The aligned lower
            // bound needs 17.36pt; a baseline-only box would reserve just 16pt.
            Assert.InRange(letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y, 17.359D, 17.361D);
        }

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
