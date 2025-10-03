using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {

        [Fact]
        public void Test_CreatingWordDocumentWithPageOrientation() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithPageOrientation.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                Assert.True(document.PageOrientation == PageOrientationValues.Portrait, "Starting page orientation should be portrait");

                document.Sections[0].PageOrientation = PageOrientationValues.Landscape;

                Assert.True(document.PageOrientation == PageOrientationValues.Landscape, "Middle page orientation should be landscape when using section 0");

                document.PageOrientation = PageOrientationValues.Portrait;

                Assert.True(document.PageOrientation == PageOrientationValues.Portrait, "Middle page orientation should be portrait when using document");

                document.AddParagraph("Test");

                document.PageOrientation = PageOrientationValues.Landscape;
                Assert.True(document.PageOrientation == PageOrientationValues.Landscape, "End page orientation should be landscape when using document");

                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs during creation is wrong.");
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong.");

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithPageOrientation.docx"))) {
                // There is only one Paragraph at the document level.
                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs during load is wrong.");
                Assert.True(document.Sections.Count == 1, "Number of sections during load is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong.");
                Assert.True(document.PageOrientation == PageOrientationValues.Landscape, "Page orientation should be landscape when using document");
                Assert.True(document.Sections[0].PageOrientation == PageOrientationValues.Landscape, "Page orientation should be landscape when using sections");
            }
        }

        [Fact]
        public void Test_GutterSettings() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithGutterSettings.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                Assert.False(document.Settings.GutterAtTop);
                document.Settings.GutterAtTop = true;
                Assert.True(document.Settings.GutterAtTop);

                Assert.False(document.RtlGutter);
                document.RtlGutter = true;
                Assert.True(document.RtlGutter);

                var second = document.AddSection();
                Assert.True(second.RtlGutter);
                second.RtlGutter = false;

                document.AddParagraph("Test");
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithGutterSettings.docx"))) {
                Assert.True(document.Settings.GutterAtTop);
                Assert.True(document.Sections[0].RtlGutter);
                Assert.False(document.Sections[1].RtlGutter);

                document.Settings.GutterAtTop = false;
                document.Sections[0].RtlGutter = false;
                document.Sections[1].RtlGutter = false;
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithGutterSettings.docx"))) {
                Assert.False(document.Settings.GutterAtTop);
                Assert.False(document.Sections[0].RtlGutter);
                Assert.False(document.Sections[1].RtlGutter);
            }
        }

        [Fact]
        public void Test_SectionParagraphsReturnEmptyWhenSectionPropertiesMissing() {
            string filePath = Path.Combine(_directoryWithFiles, "SectionParagraphsNoSectionProps.docx");
            using WordDocument document = WordDocument.Create(filePath);

            document.AddParagraph("Section zero paragraph");

            var section = document.Sections[0];
            section._sectionProperties = new SectionProperties();

            var paragraphs = section.Paragraphs;

            var texts = paragraphs.Select(p => p.Text).ToList();
            Assert.Single(paragraphs);
            Assert.Contains("Section zero paragraph", texts);
        }

        [Fact]
        public void Test_SectionParagraphsWhenBodyHasNoSectionProperties() {
            string filePath = Path.Combine(_directoryWithFiles, "SectionParagraphsDetachedProps.docx");
            using WordDocument document = WordDocument.Create(filePath);

            document.AddParagraph("First section paragraph");
            var secondSection = document.AddSection();
            secondSection.AddParagraph("Second section paragraph");

            var body = document._wordprocessingDocument.MainDocumentPart!.Document.Body!;
            foreach (var paragraph in body.Elements<Paragraph>().ToList()) {
                if (paragraph.ParagraphProperties != null) {
                    paragraph.ParagraphProperties.RemoveAllChildren<SectionProperties>();
                    if (!paragraph.ParagraphProperties.ChildElements.Any()) {
                        paragraph.ParagraphProperties.Remove();
                    }
                }
            }

            var sectionPropertiesNodes = body.Elements<SectionProperties>().ToList();
            if (sectionPropertiesNodes.Count > 1) {
                // Keep a single trailing SectionProperties to mimic documents that only store
                // section metadata at the body level, removing any earlier matches.
                foreach (var node in sectionPropertiesNodes.Take(sectionPropertiesNodes.Count - 1)) {
                    node.Remove();
                }
            }

            var bodyParagraphs = body.Elements<Paragraph>().ToList();
            var firstSectionParagraphs = document.Sections[0].Paragraphs;
            var secondSectionParagraphs = document.Sections[1].Paragraphs;

            Assert.Equal(bodyParagraphs.Count, firstSectionParagraphs.Count);
            Assert.Equal(bodyParagraphs.Count, secondSectionParagraphs.Count);
            Assert.Contains("First section paragraph", firstSectionParagraphs.Select(p => p.Text));
            Assert.Contains("Second section paragraph", secondSectionParagraphs.Select(p => p.Text));
        }
    }
}
