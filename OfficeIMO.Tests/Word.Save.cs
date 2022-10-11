using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {

        [Fact]
        public void Test_Save() {
            string filePath1 = Path.Combine(_directoryWithFiles, "FirstDocument11.docx");
            string filePath2 = Path.Combine(_directoryWithFiles, "FirstDocument12.docx");
            string filePath3 = Path.Combine(_directoryWithFiles, "FirstDocument13.docx");
            string filePath4 = Path.Combine(_directoryWithFiles, "FirstDocument14.docx");

            File.Delete(filePath1);
            File.Delete(filePath2);
            File.Delete(filePath3);
            File.Delete(filePath4);

            using (WordDocument document = WordDocument.Create()) {
                document.BuiltinDocumentProperties.Title = "This is my title";
                document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
                document.BuiltinDocumentProperties.Keywords = "word, docx, test";

                Assert.True(document.Paragraphs.Count == 0);

                document.AddParagraph("This is my test in document 1");

                Assert.True(File.Exists(filePath1) == false);

                document.Save(filePath1);

                Assert.True(File.Exists(filePath1) == true);
                Assert.True(filePath1.IsFileLocked() == true);

                Assert.True(document.Paragraphs.Count == 1);

                document.AddParagraph("This is my test in document 2");

                document.Save(filePath2);

                Assert.True(document.Paragraphs.Count == 2);

                Assert.True(File.Exists(filePath2) == true);
                Assert.True(filePath2.IsFileLocked() == true);

                document.AddParagraph("This is my test in document 3");

                Assert.True(document.Paragraphs.Count == 3);

                document.Save(filePath3);

                Assert.True(File.Exists(filePath3) == true);
                Assert.True(filePath3.IsFileLocked() == true);
            }
            Assert.True(filePath1.IsFileLocked() == false);
            Assert.True(filePath2.IsFileLocked() == false);
            Assert.True(filePath3.IsFileLocked() == false);

            using (WordDocument document = WordDocument.Load(filePath1)) {
                Assert.True(document.Paragraphs.Count == 1);
                Assert.True(filePath1.IsFileLocked() == true);
            }
            using (WordDocument document = WordDocument.Load(filePath2)) {
                Assert.True(document.Paragraphs.Count == 2);
                Assert.True(filePath2.IsFileLocked() == true);
            }
            using (WordDocument document = WordDocument.Load(filePath3)) {
                Assert.True(filePath3.IsFileLocked() == true);

                Assert.True(document.Paragraphs.Count == 3);
                document.AddParagraph("More paragraphs!");
                Assert.True(document.Paragraphs.Count == 4);
                document.Save(filePath4);
            }

            Assert.True(filePath1.IsFileLocked() == false);
            Assert.True(filePath2.IsFileLocked() == false);
            Assert.True(filePath3.IsFileLocked() == false);
            Assert.True(filePath4.IsFileLocked() == false);

            using (WordDocument document = WordDocument.Load(filePath3)) {
                Assert.True(document.Paragraphs.Count == 3);
                Assert.True(filePath3.IsFileLocked() == true);
            }
            using (WordDocument document = WordDocument.Load(filePath4)) {
                Assert.True(document.Paragraphs.Count == 4);
                Assert.True(filePath4.IsFileLocked() == true);
            }

            Assert.True(filePath1.IsFileLocked() == false);
            Assert.True(filePath2.IsFileLocked() == false);
            Assert.True(filePath3.IsFileLocked() == false);
            Assert.True(filePath4.IsFileLocked() == false);
        }

        [Fact]
        public void Test_Dispose() {
            string filePath1 = Path.Combine(_directoryWithFiles, "DisposeTesting.docx");
            File.Delete(filePath1);

            Assert.True(File.Exists(filePath1) == false);

            WordDocument document = WordDocument.Create(filePath1);
            document.BuiltinDocumentProperties.Title = "This is my title";
            document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
            document.BuiltinDocumentProperties.Keywords = "word, docx, test";

            document.AddParagraph("This is my test");

            Assert.True(filePath1.IsFileLocked() == true);

            document.Save();

            Assert.True(filePath1.IsFileLocked() == true);

            document.Dispose();

            Assert.True(filePath1.IsFileLocked() == false);

            Assert.True(File.Exists(filePath1) == true);
        }

        [Fact]
        public void Test_SaveToStream()
        {
            var document = WordDocument.Create();
            document.AddParagraph("Hello world!");

            using (var outputStream = new MemoryStream())
            {
                document.Save(outputStream);

                var resultDoc = WordDocument.Load(outputStream);

                var expectedCount = 1;
                var expectedText = "Hello world!";

                Assert.Equal(resultDoc.Paragraphs.Count, expectedCount);
                Assert.Equal(resultDoc.Paragraphs[0].Text, expectedText);
            }
        }
    }
}
