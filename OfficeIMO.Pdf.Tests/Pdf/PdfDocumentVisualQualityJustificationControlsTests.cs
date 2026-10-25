using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public partial class PdfDocumentVisualQualityTests {
        [Theory]
        [InlineData("ab\u200E", false)]
        [InlineData("ab\u200E", true)]
        [InlineData("\u200Eab", false)]
        [InlineData("\u200Eab", true)]
        [InlineData("a\u200Eb", false)]
        [InlineData("a\u200Eb", true)]
        [InlineData("a\u200Eb\u200E", false)]
        [InlineData("a\u200Eb\u200E", true)]
        public void CharacterJustificationRetainedControlsDoNotCreateExpansionOpportunities(string text, bool wordSpacing) {
            var options = JustificationPolicyOptions().UseFontFamily(new PdfEmbeddedFontFamily("Control Distribution",
                ManagedTextShapingTestAssets.CreateCharacterDistributionControlFont()))
                .SetTextShapingProvider(OfficeManagedTextShapingProvider.Instance);
            byte[] Render(string value) => PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text(value),
                PdfAlign.Justify, style: new PdfParagraphStyle {
                    JustificationMode = PdfJustificationMode.Characters, JustifyLastLine = true,
                    UseWordJustificationSpacing = wordSpacing
                }).ToBytes();

            byte[] control = Render("ab"), actual = Render(text);
            var expectedLetters = ReadJustificationPolicyLetters(control).Where(letter => letter.Value is "a" or "b").ToArray();
            var actualLetters = ReadJustificationPolicyLetters(actual).Where(letter => letter.Value is "a" or "b").ToArray();
            Assert.Equal(2, expectedLetters.Length);
            Assert.Equal(2, actualLetters.Length);
            for (int index = 0; index < actualLetters.Length; index++) {
                Assert.Equal(expectedLetters[index].Value, actualLetters[index].Value);
                Assert.Equal(expectedLetters[index].StartBaseLine.X, actualLetters[index].StartBaseLine.X, 3);
                Assert.Equal(expectedLetters[index].EndBaseLine.X, actualLetters[index].EndBaseLine.X, 3);
            }
            Assert.Equal(270D, actualLetters.Last().EndBaseLine.X, 3);
            Assert.Equal(text, PdfReadDocument.Open(actual).ExtractText().Trim());
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void CharacterJustificationSingleVisibleClusterWithControlsRetainsNaturalWidth(bool wordSpacing) {
            var options = JustificationPolicyOptions().UseFontFamily(new PdfEmbeddedFontFamily("Control Distribution",
                ManagedTextShapingTestAssets.CreateCharacterDistributionControlFont()))
                .SetTextShapingProvider(OfficeManagedTextShapingProvider.Instance);
            byte[] actual = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text("a\u200E"),
                PdfAlign.Justify, style: new PdfParagraphStyle {
                    JustificationMode = PdfJustificationMode.Characters, JustifyLastLine = true,
                    UseWordJustificationSpacing = wordSpacing
                }).ToBytes();
            var letter = Assert.Single(ReadJustificationPolicyLetters(actual).Where(item => item.Value == "a"));
            Assert.Equal(30D, letter.StartBaseLine.X, 3);
            Assert.Equal(36D, letter.EndBaseLine.X, 3);
            byte[] natural = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text("a\u200E")).ToBytes();
            // A zero-advance control can still require a TJ correction to its
            // nominal encoded width. Distribution must not add another one.
            Assert.Equal(JustificationTextAdjustments(natural), JustificationTextAdjustments(actual));
        }
    }
}

namespace OfficeIMO.TestAssets {
    internal static partial class ManagedTextShapingTestAssets {
        internal static byte[] CreateCharacterDistributionControlFont() =>
            CreateFontFromCmap(CreateDistinctFormat12Cmap(new[] { 32, (int)'a', 'b', 0x200E }),
                glyphCount: 5, blankGlyphIds: new[] { 1, 4 });
    }
}
