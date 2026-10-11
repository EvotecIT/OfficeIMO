using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;
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

        [Theory]
        [InlineData(false, false)]
        [InlineData(false, true)]
        [InlineData(true, false)]
        [InlineData(true, true)]
        public void SaveAsPdf_OrdinaryRunningStoriesReserveTheOppositePhysicalFrame(bool footer, bool oppositeTable) {
            using WordDocument document = CreateJoinedParagraphDocument();
            document.Sections[0].PageSettings.PageSize = WordPageSize.Letter;
            document.Sections[0].Margins.Top = document.Sections[0].Margins.Bottom = 1440;
            document.Sections[0].Margins.HeaderDistance = document.Sections[0].Margins.FooterDistance = 720;
            WordHeaderFooter story = footer ? document.FooterDefaultOrCreate : document.HeaderDefaultOrCreate;
            WordParagraph paragraph = story.AddParagraph("STORY");
            paragraph._paragraph.ParagraphProperties = new W.ParagraphProperties(new W.SpacingBetweenLines {
                Before = "0", After = "0", Line = "14600", LineRule = W.LineSpacingRuleValues.Exact });
            if (oppositeTable) {
                WordHeaderFooter opposite = footer ? document.HeaderDefaultOrCreate : document.FooterDefaultOrCreate;
                opposite.AddTable(1, 1).Rows[0].Cells[0].AddParagraph("REQUIRED");
            }
            document.AddParagraph("BODY");

            using var pdf = OpenJoinedParagraphPdf(document);
            Assert.Equal(1, pdf.NumberOfPages);
            Assert.Contains("BODY", pdf.GetPage(1).Text);
            Assert.Contains("STORY", pdf.GetPage(1).Text);
            if (oppositeTable) Assert.Contains("REQUIRED", pdf.GetPage(1).Text);
        }

        [Theory]
        [InlineData(false, 1)]
        [InlineData(false, 10)]
        [InlineData(true, 1)]
        [InlineData(true, 10)]
        public void SaveAsPdf_OrdinaryRunningStoryFitUsesTheActualPageNumber(bool footer, int start) {
            using WordDocument document = CreateJoinedParagraphDocument();
            document.Sections[0].PageSettings.PageSize = WordPageSize.Letter;
            document.Sections[0]._sectionProperties.AddChild(new W.PageNumberType { Start = start }, true);
            WordHeaderFooter story = footer ? document.FooterDefaultOrCreate : document.HeaderDefaultOrCreate;
            WordParagraph paragraph = story.AddParagraph();
            paragraph.AddField(WordFieldType.Page);
            foreach (W.Run run in paragraph._paragraph.Descendants<W.Run>()) {
                run.RunProperties = new W.RunProperties(new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                    new W.FontSize { Val = "900" }, new W.FontSizeComplexScript { Val = "900" });
            }
            document.AddParagraph("BODY");

            using var pdf = OpenJoinedParagraphPdf(document);
            Assert.Equal(1, pdf.NumberOfPages);
            Assert.Contains("BODY", pdf.GetPage(1).Text);
            Assert.Contains(start.ToString(), pdf.GetPage(1).Text);
        }

        [Fact]
        public void SaveAsPdf_OrdinaryRunningHeaderRetainsRoomForTheFirstBodyLine() {
            using WordDocument document = CreateJoinedParagraphDocument();
            document.Sections[0].PageSettings.PageSize = WordPageSize.Letter;
            document.Sections[0].Margins.Bottom = 1440;
            document.Sections[0].Margins.HeaderDistance = 720;
            WordParagraph header = document.HeaderDefaultOrCreate.AddParagraph("HEADER");
            header._paragraph.ParagraphProperties = new W.ParagraphProperties(new W.SpacingBetweenLines {
                Before = "0", After = "0", Line = "13000", LineRule = W.LineSpacingRuleValues.Exact });
            document.AddParagraph("BODY").FontSize = 40;

            using var pdf = OpenJoinedParagraphPdf(document);
            Assert.Equal(1, pdf.NumberOfPages);
            Assert.Contains("HEADER", pdf.GetPage(1).Text);
            Assert.Contains("BODY", pdf.GetPage(1).Text);
        }
    }
}
