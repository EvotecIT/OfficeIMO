using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void InspectionSnapshotResolvesCaseDistinctCustomStyleNames() {
            using var stream = new MemoryStream();
            using WordDocument document = WordDocument.Create(stream);
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(
                new Style(new StyleName { Val = "Uppercase custom name" }) {
                    Type = StyleValues.Paragraph,
                    StyleId = "InspectionCustom",
                    CustomStyle = true,
                },
                new Style(new StyleName { Val = "Lowercase custom name" }) {
                    Type = StyleValues.Paragraph,
                    StyleId = "inspectioncustom",
                    CustomStyle = true,
                });

            document.AddParagraph("First").SetStyleId("inspectioncustom");
            document.AddParagraph("Second").SetStyleId("InspectionCustom");

            WordDocumentInfo snapshot = document.Inspect();
            WordParagraphInfo[] paragraphs = snapshot.Sections
                .SelectMany(section => section.Elements.OfType<WordParagraphInfo>())
                .ToArray();

            Assert.Equal(2, paragraphs.Length);
            Assert.Equal(new[] { "Lowercase custom name", "Uppercase custom name" },
                paragraphs.Select(paragraph => paragraph.StyleName));
        }
    }
}
