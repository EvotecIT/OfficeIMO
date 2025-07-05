using OfficeIMO.Word;
using Xunit;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AddingStructuredDocumentTag() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithContentControl.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var sdt = document.AddStructuredDocumentTag("Hello world", "Alias1");

                Assert.True(document.StructuredDocumentTags.Count == 1);
                Assert.True(document.ParagraphsStructuredDocumentTags.Count == 1);
                Assert.Equal("Hello world", sdt.Text);

                document.Save(false);

                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.StructuredDocumentTags.Count == 1);
                Assert.Equal("Hello world", document.StructuredDocumentTags[0].Text);

                document.StructuredDocumentTags[0].Text = "Changed";
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal("Changed", document.StructuredDocumentTags[0].Text);
            }
        }

        [Fact]
        public void Test_StructuredDocumentTagWithTag() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithContentControlTag.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var sdt = document.AddStructuredDocumentTag("Hello", "Alias1", "Tag1");

                Assert.Equal("Tag1", sdt.Tag);
                document.Save(false);
                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var loaded = document.GetStructuredDocumentTagByTag("Tag1");
                Assert.NotNull(loaded);
                Assert.Equal("Hello", loaded.Text);

                loaded.Text = "Updated";
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal("Updated", document.StructuredDocumentTags[0].Text);
                Assert.Equal("Tag1", document.StructuredDocumentTags[0].Tag);
            }
        }

        [Fact]
        public void Test_StructuredDocumentTagGetByAlias() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithContentControlAlias.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var sdt = document.AddStructuredDocumentTag("Hello", "Alias100", "Tag100");

                Assert.NotNull(document.GetStructuredDocumentTagByAlias("Alias100"));
                document.Save(false);
                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var alias = document.GetStructuredDocumentTagByAlias("Alias100");
                Assert.NotNull(alias);
                alias.Text = "Updated";
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal("Updated", document.StructuredDocumentTags[0].Text);
                Assert.Equal("Tag100", document.StructuredDocumentTags[0].Tag);
            }
        }

        [Fact]
        public void Test_SettingTextOnEmptyStructuredDocumentTag() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithEmptySdt.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var sdtRun = new SdtRun();

                var properties = new SdtProperties();
                properties.Append(new SdtId() { Val = new DocumentFormat.OpenXml.Int32Value(new System.Random().Next(1, int.MaxValue)) });
                sdtRun.Append(properties);
                sdtRun.Append(new SdtContentRun());

                var paragraph = new Paragraph(sdtRun);
                var wordParagraph = new WordParagraph(document, paragraph, sdtRun);
                document.AddParagraph(wordParagraph);

                var sdt = wordParagraph.StructuredDocumentTag;
                Assert.Null(sdt.Text);

                sdt.Text = "New text";

                Assert.Equal("New text", sdt.Text);
                Assert.NotNull(sdtRun.SdtContentRun.GetFirstChild<Run>());
                Assert.Equal("New text", sdtRun.SdtContentRun.GetFirstChild<Run>().GetFirstChild<Text>().Text);

                document.Save(false);
                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.StructuredDocumentTags);
                Assert.Equal("New text", document.StructuredDocumentTags[0].Text);
            }
        }
    }
}
