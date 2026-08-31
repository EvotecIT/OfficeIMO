using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void InspectionSnapshotResolvesCustomStyleNamesFromOneCaseInsensitiveCatalog() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(
                new Style(new StyleName { Val = "First custom name" }) {
                    Type = StyleValues.Paragraph,
                    StyleId = "InspectionCustom",
                    CustomStyle = true,
                },
                new Style(new StyleName { Val = "Duplicate custom name" }) {
                    Type = StyleValues.Paragraph,
                    StyleId = "inspectioncustom",
                    CustomStyle = true,
                });

            document.AddParagraph("First").SetStyleId("inspectioncustom");
            document.AddParagraph("Second").SetStyleId("InspectionCustom");

            WordDocumentSnapshot snapshot = document.CreateInspectionSnapshot();
            WordParagraphSnapshot[] paragraphs = snapshot.Sections
                .SelectMany(section => section.Elements.OfType<WordParagraphSnapshot>())
                .ToArray();

            Assert.Equal(2, paragraphs.Length);
            Assert.All(paragraphs, paragraph => Assert.Equal("First custom name", paragraph.StyleName));
        }
    }
}
