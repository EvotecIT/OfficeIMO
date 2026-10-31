using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using UglyToad.PdfPig.Content;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf {
    public sealed class PdfNaturalFontLineAlignmentTests {
        [Fact]
        public void AscentOnlyStrutKeepsTheSameGridBaselineAsVisibleAscentOnlyText() {
            byte[] font = ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(
                1000, -500, 0, 500, Enumerable.Range(32, 95).ToArray());
            byte[] Render(bool strut) {
                var options = new PdfCreateOptions {
                    PageWidth = global::OfficeIMO.OfficeLength.Points(300D), PageHeight = global::OfficeIMO.OfficeLength.Points(300D),
                    MarginTop = global::OfficeIMO.OfficeLength.Points(30D), MarginBottom = global::OfficeIMO.OfficeLength.Points(30D), MarginLeft = global::OfficeIMO.OfficeLength.Points(30D), MarginRight = global::OfficeIMO.OfficeLength.Points(30D)
                };
                options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Strut grid metrics", font));
                var ascentOnly = PdfTextRun.Normal(strut ? string.Empty : "X", fontSize: global::OfficeIMO.OfficeLength.Points(40D),
                    fontFamily: "Strut grid metrics");
                ascentOnly.IsLineBoxStrut = strut;
                ascentOnly.ContributesLineBoxDescent = false;
                var style = new PdfParagraphStyle {
                    FontSize = global::OfficeIMO.OfficeLength.Points(11D), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D),
                    LineSpacing = PdfLineSpacing.Multiple(1.5D).WithFontLineBoxBaseline(1.5D).WithLineGrid(18D)
                };
                return PdfDocument.Create(options)
                    .Paragraph(paragraph => paragraph.Runs(new[] {
                        PdfTextRun.Normal("A", fontSize: global::OfficeIMO.OfficeLength.Points(11D), fontFamily: "Strut grid metrics"), ascentOnly
                    }), style: style)
                    .Paragraph(paragraph => paragraph.Runs(new[] {
                        PdfTextRun.Normal("B", fontSize: global::OfficeIMO.OfficeLength.Points(11D), fontFamily: "Strut grid metrics")
                    }), style: style).ToBytes();
            }
            using PdfPigDocument actual = PdfPigDocument.Open(Render(strut: true));
            using PdfPigDocument visibleReference = PdfPigDocument.Open(Render(strut: false));
            Letter first = Assert.Single(actual.GetPage(1).Letters, letter => letter.Value == "A");
            Letter next = Assert.Single(actual.GetPage(1).Letters, letter => letter.Value == "B");
            Letter expectedFirst = Assert.Single(visibleReference.GetPage(1).Letters, letter => letter.Value == "A");
            Letter expectedNext = Assert.Single(visibleReference.GetPage(1).Letters, letter => letter.Value == "B");
            // The first line occupies 40pt ascent + 5.5pt body descent, snapped
            // to 54pt. Its baseline is 44.25pt below the top; the next is 11.75pt.
            Assert.Equal(21.5D, first.StartBaseLine.Y - next.StartBaseLine.Y, 3);
            Assert.Equal(expectedFirst.StartBaseLine.Y, first.StartBaseLine.Y, 3);
            Assert.Equal(expectedNext.StartBaseLine.Y, next.StartBaseLine.Y, 3);
            Assert.DoesNotContain(actual.GetPage(1).Letters, letter => letter.Value == "X");
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void AlignedNaturalLinesReserveTheirCombinedFontBox(bool minimum) {
            PdfCreateOptions options = new PdfCreateOptions();
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Natural line bounds", ManagedTextShapingTestAssets.CreateFontWithDeviceMetrics()));
            PdfLineSpacing spacing = (minimum ? PdfLineSpacing.AtLeast(global::OfficeIMO.OfficeLength.Points(1), 1D) : PdfLineSpacing.Multiple(1D))
                .WithFontLineBoxBaseline(1D).WithFontAlignment(PdfFontLineAlignment.Top, .12D);
            PdfDocument document = PdfDocument.Create(options);
            document.Compose(builder => builder.Content(content => content.Paragraph(paragraph => paragraph
                .FontFamily("Natural line bounds").FontSize(global::OfficeIMO.OfficeLength.Points(8)).Text("A").FontSize(global::OfficeIMO.OfficeLength.Points(16)).Text("B")
                .Text("\n").FontSize(global::OfficeIMO.OfficeLength.Points(8)).Text("A").FontSize(global::OfficeIMO.OfficeLength.Points(16)).Text("B"),
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
            PdfCreateOptions options = new PdfCreateOptions();
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Natural device metrics", ManagedTextShapingTestAssets.CreateFontWithDeviceMetrics()));
            PdfLineSpacing spacing = (minimum ? PdfLineSpacing.AtLeast(global::OfficeIMO.OfficeLength.Points(30), 1.2D) : PdfLineSpacing.Multiple(1.2D))
                .WithFontLineBoxBaseline(1.2D).WithFontAlignment(alignment, .12D);
            PdfDocument document = PdfDocument.Create(options);
            document.Paragraph(paragraph => paragraph.FontFamily("Natural device metrics").FontSize(global::OfficeIMO.OfficeLength.Points(8)).Text("A")
                .FontSize(global::OfficeIMO.OfficeLength.Points(16)).Text("B"), style: new PdfParagraphStyle { LineSpacing = spacing });
            using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes());
            IReadOnlyList<Letter> letters = pdf.GetPage(1).Letters;
            Letter first = Assert.Single(letters, letter => letter.Value == "A");
            Letter second = Assert.Single(letters, letter => letter.Value == "B");
            Assert.InRange(first.StartBaseLine.Y - second.StartBaseLine.Y, difference - .002D, difference + .002D);
        }
    }
}
