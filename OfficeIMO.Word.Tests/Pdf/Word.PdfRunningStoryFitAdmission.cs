using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Theory]
        [InlineData(false, false)]
        [InlineData(false, true)]
        [InlineData(true, false)]
        [InlineData(true, true)]
        public void SaveAsPdf_OrdinaryRunningStoriesRequireTheCompleteWrappedHeightToFit(bool footer, bool largeRun) {
            using WordDocument document = CreateJoinedParagraphDocument();
            WordHeaderFooter story = footer ? document.FooterDefaultOrCreate : document.HeaderDefaultOrCreate;
            WordParagraph paragraph = story.AddParagraph("Small ");
            paragraph.FontSize = 8;
            if (largeRun) {
                paragraph.AddText("Large").FontSize = 800;
            } else {
                paragraph.AddText(string.Concat(Enumerable.Repeat("ALPHA BETA ", 600)));
            }
            document.AddParagraph("BODY");

            using var pdf = OpenJoinedParagraphPdf(document);
            Assert.Equal(1, pdf.NumberOfPages);
            Assert.Contains("BODY", pdf.GetPage(1).Text);
        }
    }
}
