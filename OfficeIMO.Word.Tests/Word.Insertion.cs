using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_InsertParagraphAt() {
            string filePath = Path.Combine(_directoryWithFiles, "InsertParagraphAt.docx");
            using (var document = WordDocument.Create(filePath)) {
                document.AddParagraph("Paragraph 1");
                document.AddParagraph("Paragraph 2");
                document.AddParagraph("Paragraph 3");

                var newParagraph = new WordParagraph(document, true, false) { Text = "Inserted" };
                document.InsertParagraphAt(1, newParagraph);

                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var body = document._document.Body!;
                Assert.Equal("Paragraph 1", body.Elements<Paragraph>().ElementAt(0).InnerText);
                Assert.Equal("Inserted", body.Elements<Paragraph>().ElementAt(1).InnerText);
                Assert.Equal("Paragraph 2", body.Elements<Paragraph>().ElementAt(2).InnerText);
            }
        }

        [Fact]
        public void Test_InsertTableAfter() {
            string filePath = Path.Combine(_directoryWithFiles, "InsertTableAfter.docx");
            using (var document = WordDocument.Create(filePath)) {
                var p1 = document.AddParagraph("Before");
                document.AddParagraph("After");

                var table = document.CreateTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Test";

                document.InsertTableAfter(p1, table);
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Single(document.Tables);
                Assert.Equal("Test", document.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);

                var body = document._document.Body!;
                Assert.IsType<Paragraph>(body.ChildElements[0]);
                Assert.IsType<Table>(body.ChildElements[1]);
                Assert.IsType<Paragraph>(body.ChildElements[2]);
            }
        }
    }
}
