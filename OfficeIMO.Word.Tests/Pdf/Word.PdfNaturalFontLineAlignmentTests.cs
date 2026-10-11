using OfficeIMO.Word;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Theory]
        [InlineData(false, false, WordLineSpacingRule.Auto)]
        [InlineData(false, true, WordLineSpacingRule.Auto)]
        [InlineData(true, false, WordLineSpacingRule.Auto)]
        [InlineData(true, true, WordLineSpacingRule.Auto)]
        [InlineData(false, false, WordLineSpacingRule.AtLeast)]
        [InlineData(false, true, WordLineSpacingRule.AtLeast)]
        [InlineData(true, false, WordLineSpacingRule.AtLeast)]
        [InlineData(true, true, WordLineSpacingRule.AtLeast)]
        public void SaveAsPdf_NaturalAndMinimumLineAlignmentRetainsMixedRunMetrics(bool nativeDoc, bool table,
            WordLineSpacingRule rule) {
            using PdfPigDocument pdf = PdfPigDocument.Open(CreateLineAlignmentPdf("direct", "top", nativeDoc, table, rule));
            AssertLineAlignmentDifference(pdf, 7.08D);
        }
    }
}
